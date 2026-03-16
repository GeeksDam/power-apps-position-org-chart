import { IInputs, IOutputs } from "./generated/ManifestTypes";

// ─── Data model ──────────────────────────────────────────────────────────────

interface PositionRecord {
    id: string;
    name: string;
    parentId: string | null;
}

interface TreeNode {
    id: string;
    name: string;
    children: TreeNode[];
    isSelected: boolean;
    isAncestor: boolean;
}

// ─── Layout constants ─────────────────────────────────────────────────────────

const NODE_W = 80;    // node box width  (px)
const NODE_H = 24;    // node box height (px)
const H_GAP  = 12;    // horizontal gap between sibling subtrees (px)
const V_GAP  = 32;    // vertical gap between tree levels (px)
const PAD    = 15;    // canvas outer padding (px)



// ─── Control class ────────────────────────────────────────────────────────────

export class positionchart implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _container!: HTMLDivElement;
    private _context!: ComponentFramework.Context<IInputs>;
    private _positionsMap: Map<string, PositionRecord> = new Map<string, PositionRecord>();
    private _childrenMap: Map<string, string[]>        = new Map<string, string[]>();
    private _textInput!: HTMLInputElement;
    private _chartContainer!: HTMLDivElement;
    private _statusLabel!: HTMLDivElement;

    // Track last config so updateView can detect property changes
    private _lastTable     = "";
    private _lastNameCol   = "";
    private _lastParentCol = "";

    constructor() { /* empty */ }

    // ──────────────────────────────────────────────────────────────────────────
    // Lifecycle
    // ──────────────────────────────────────────────────────────────────────────

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._container = container;
        this._context   = context;
        this._buildUI();
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._context = context;
        const cfg = this._getConfig();

        // Pre-fill text input from bound field if provided and input is empty
        const boundName = (context.parameters.defaultPositionName.raw ?? "").trim();
        if (boundName && !this._textInput.value) {
            this._textInput.value = boundName;
        }

        // Reload whenever table/column config changes
        if (
            cfg.table     !== this._lastTable     ||
            cfg.nameCol   !== this._lastNameCol   ||
            cfg.parentCol !== this._lastParentCol
        ) {
            this._lastTable     = cfg.table;
            this._lastNameCol   = cfg.nameCol;
            this._lastParentCol = cfg.parentCol;
            if (cfg.table && cfg.nameCol && cfg.parentCol) {
                void this._loadPositions();
            }
        }
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void {
        this._container.innerHTML = "";
    }

    // ──────────────────────────────────────────────────────────────────────────
    // UI construction
    // ──────────────────────────────────────────────────────────────────────────

    private _buildUI(): void {
        this._container.className = "pos-chart-root";

        // ── toolbar ──────────────────────────────────────────────────────────
        const toolbar = document.createElement("div");
        toolbar.className = "pos-chart-toolbar";

        const labelEl = document.createElement("label");
        labelEl.textContent = "Position Name:";
        labelEl.className = "pos-chart-label";
        labelEl.setAttribute("for", "pos-chart-input");

        this._textInput = document.createElement("input");
        this._textInput.type        = "text";
        this._textInput.id          = "pos-chart-input";
        this._textInput.className   = "pos-chart-select";
        this._textInput.placeholder = "Type a position name…";
        this._textInput.addEventListener("keydown", (e) => {
            if (e.key === "Enter") this._buildChart();
        });

        const btn = document.createElement("button");
        btn.textContent = "Build Chart";
        btn.className   = "pos-chart-btn";
        btn.type        = "button";
        btn.addEventListener("click", () => this._buildChart());

        toolbar.appendChild(labelEl);
        toolbar.appendChild(this._textInput);
        toolbar.appendChild(btn);

        // ── status bar ───────────────────────────────────────────────────────
        this._statusLabel = document.createElement("div");
        this._statusLabel.className = "pos-chart-status pos-chart-status--info";

        // ── chart canvas ─────────────────────────────────────────────────────
        this._chartContainer = document.createElement("div");
        this._chartContainer.className = "pos-chart-area";

        this._container.appendChild(toolbar);
        this._container.appendChild(this._statusLabel);
        this._container.appendChild(this._chartContainer);
    }

    // ──────────────────────────────────────────────────────────────────────────
    // Config helpers
    // ──────────────────────────────────────────────────────────────────────────

    private _getConfig(): { table: string; nameCol: string; parentCol: string } {
        const table     = (this._context.parameters.tableLogicalName.raw  ?? "").trim();
        const nameCol   = (this._context.parameters.nameColumn.raw         ?? "").trim();
        const parentCol = (this._context.parameters.parentLookupColumn.raw ?? "").trim();
        return { table, nameCol, parentCol };
    }

    // ──────────────────────────────────────────────────────────────────────────
    // Data loading
    // ──────────────────────────────────────────────────────────────────────────

    private async _loadPositions(): Promise<void> {
        const { table, nameCol, parentCol } = this._getConfig();

        // Validate logical names to prevent OData query injection
        if (!this._isValidLogicalName(table) || !this._isValidLogicalName(nameCol) || !this._isValidLogicalName(parentCol)) {
            this._setStatus("Configuration error: table and column names must contain only lowercase letters, digits, and underscores.", "error");
            return;
        }

        const parentValCol = `_${parentCol}_value`;
        const idCol        = `${table}id`;

        this._setStatus("Loading positions from Dataverse…", "info");
        this._positionsMap.clear();
        this._childrenMap.clear();

        try {
            const options = `?$select=${nameCol},${parentValCol}&$orderby=${nameCol} asc`;
            await this._fetchAllPages(table, options, idCol, nameCol, parentValCol);

            this._setStatus(
                `${this._positionsMap.size} position(s) loaded. Type a position name and click "Build Chart".`,
                "info"
            );
        } catch (err) {
            this._setStatus(`Error loading positions: ${(err as Error).message}`, "error");
        }
    }

    /** Returns true only for valid Dataverse logical names (lowercase letters, digits, underscores, max 100 chars). */
    private _isValidLogicalName(name: string): boolean {
        return /^[a-z_][a-z0-9_]{0,99}$/.test(name);
    }

    /** Page through all Dataverse records handling OData @odata.nextLink paging. */
    private async _fetchAllPages(
        table: string,
        options: string,
        idCol: string,
        nameCol: string,
        parentValCol: string
    ): Promise<void> {
        let nextOptions: string | undefined = options;
        let pageCount = 0;
        const MAX_PAGES = 50; // safety cap — 50 × 5 000 = 250 000 records max

        while (nextOptions) {
            if (++pageCount > MAX_PAGES) {
                this._setStatus(`Loaded first ${this._positionsMap.size} positions (page limit reached).`, "warning");
                break;
            }
            const result = await this._context.webAPI.retrieveMultipleRecords(
                table,
                nextOptions,
                5000
            );

            for (const record of result.entities) {
                const id       = record[idCol] as string;
                const name     = (record[nameCol] as string) ?? "(Unnamed)";
                const parentId = (record[parentValCol] as string) ?? null;

                this._positionsMap.set(id, { id, name, parentId });

                if (!this._childrenMap.has(id)) {
                    this._childrenMap.set(id, []);
                }
                if (parentId) {
                    if (!this._childrenMap.has(parentId)) {
                        this._childrenMap.set(parentId, []);
                    }
                    this._childrenMap.get(parentId)!.push(id);
                }
            }

            // Advance to next page if present
            if (result.nextLink) {
                const qIdx = result.nextLink.indexOf("?");
                nextOptions = qIdx >= 0 ? result.nextLink.substring(qIdx) : undefined;
            } else {
                nextOptions = undefined;
            }
        }
    }

    // ──────────────────────────────────────────────────────────────────────────
    // Chart building
    // ──────────────────────────────────────────────────────────────────────────

    private _buildChart(): void {
        const typedName = this._textInput.value.trim();
        if (!typedName) {
            this._setStatus("Please enter a position name before building the chart.", "warning");
            return;
        }

        const lowerTyped = typedName.toLowerCase();
        const selected = Array.from(this._positionsMap.values())
            .find(p => p.name.toLowerCase() === lowerTyped);
        if (!selected) {
            this._setStatus(`No position found with name "${typedName}". Check spelling and try again.`, "error");
            return;
        }
        const selectedId = selected.id;

        // ── 1. Collect ancestor IDs (root-first order) ────────────────────────
        const ancestorIds: string[] = [];
        let curId: string | null    = selected.parentId;
        const visited               = new Set<string>();

        while (curId && !visited.has(curId)) {
            visited.add(curId);
            ancestorIds.unshift(curId);
            const parent = this._positionsMap.get(curId);
            curId = parent ? parent.parentId : null;
        }

        // ── 2. Build full display tree ────────────────────────────────────────
        //    Ancestor chain ──► selected ──► all descendants
        const selectedNode = this._buildDescendants(selectedId, true);

        let root: TreeNode;
        if (ancestorIds.length === 0) {
            root = selectedNode;
        } else {
            root = {
                id:         ancestorIds[0],
                name:       this._positionsMap.get(ancestorIds[0])!.name,
                children:   [],
                isSelected: false,
                isAncestor: true,
            };
            let ptr = root;
            for (let i = 1; i < ancestorIds.length; i++) {
                const node: TreeNode = {
                    id:         ancestorIds[i],
                    name:       this._positionsMap.get(ancestorIds[i])!.name,
                    children:   [],
                    isSelected: false,
                    isAncestor: true,
                };
                ptr.children.push(node);
                ptr = node;
            }
            ptr.children.push(selectedNode);
        }

        // ── 3. Render ─────────────────────────────────────────────────────────
        this._renderChart(root);
        this._setStatus(
            `Org chart for "${selected.name}" — ${ancestorIds.length} ancestor(s) above, ` +
            `${this._countDescendants(selectedNode) - 1} descendant position(s) below.`,
            "info"
        );
    }

    /** Recursively build a TreeNode subtree from a given position downward.
     *  The visited set prevents infinite recursion if Dataverse data contains circular parent references.
     */
    private _buildDescendants(nodeId: string, isSelected: boolean, visited: Set<string> = new Set<string>()): TreeNode {
        visited.add(nodeId);
        const pos           = this._positionsMap.get(nodeId)!;
        const safeChildIds  = (this._childrenMap.get(nodeId) ?? []).filter(cid => !visited.has(cid));
        return {
            id:         nodeId,
            name:       pos.name,
            children:   safeChildIds.map(cid => this._buildDescendants(cid, false, visited)),
            isSelected,
            isAncestor: false,
        };
    }

    private _countDescendants(node: TreeNode): number {
        return 1 + node.children.reduce((sum, c) => sum + this._countDescendants(c), 0);
    }

    // ──────────────────────────────────────────────────────────────────────────
    // SVG-based chart renderer
    // ──────────────────────────────────────────────────────────────────────────

    private _renderChart(root: TreeNode): void {
        // ── Step 1: Calculate subtree widths (post-order) ─────────────────────
        const subtreeW = new Map<TreeNode, number>();

        const calcWidth = (node: TreeNode): number => {
            if (node.children.length === 0) {
                subtreeW.set(node, NODE_W);
                return NODE_W;
            }
            const total = node.children.reduce(
                (sum, c) => sum + calcWidth(c) + H_GAP, 0
            ) - H_GAP;
            const w = Math.max(NODE_W, total);
            subtreeW.set(node, w);
            return w;
        };
        calcWidth(root);

        // ── Step 2: Assign (x, y) coordinates ────────────────────────────────
        interface PosEntry { x: number; y: number; parent: TreeNode | null; }
        const positions = new Map<TreeNode, PosEntry>();

        const assignPos = (
            node: TreeNode,
            left: number,
            depth: number,
            parent: TreeNode | null
        ): void => {
            const sw = subtreeW.get(node)!;
            const x  = left + (sw - NODE_W) / 2;
            const y  = depth * (NODE_H + V_GAP) + PAD;
            positions.set(node, { x, y, parent });

            let childLeft = left;
            for (const child of node.children) {
                assignPos(child, childLeft, depth + 1, node);
                childLeft += subtreeW.get(child)! + H_GAP;
            }
        };
        assignPos(root, PAD, 0, null);

        // ── Step 3: Compute canvas dimensions ─────────────────────────────────
        let canvasW = 0;
        let canvasH = 0;
        for (const p of positions.values()) {
            canvasW = Math.max(canvasW, p.x + NODE_W + PAD);
            canvasH = Math.max(canvasH, p.y + NODE_H + PAD);
        }

        // ── Step 4: Build DOM ─────────────────────────────────────────────────
        this._chartContainer.innerHTML = "";

        const wrapper = document.createElement("div");
        wrapper.style.position = "relative";
        wrapper.style.width    = `${canvasW}px`;
        wrapper.style.height   = `${canvasH}px`;

        // SVG for connector lines (sits beneath node divs)
        const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        svg.setAttribute("width",  canvasW.toString());
        svg.setAttribute("height", canvasH.toString());
        svg.style.position = "absolute";
        svg.style.top  = "0";
        svg.style.left = "0";
        svg.style.overflow = "visible";
        wrapper.appendChild(svg);

        // ── Step 5: Draw connector lines ──────────────────────────────────────
        //    Pattern: parent-bottom → trunk → horizontal bar → drops → child-top
        for (const [node, pos] of positions.entries()) {
            if (node.children.length === 0) continue;

            const px = pos.x + NODE_W / 2;   // parent center-x
            const py = pos.y + NODE_H;        // parent bottom-y
            // Trunk drops 40% of V_GAP before the horizontal bar
            const midY = py + V_GAP * 0.4;

            this._svgLine(svg, px, py, px, midY);   // vertical trunk

            if (node.children.length === 1) {
                // Single child: straight line all the way to child top
                const cp = positions.get(node.children[0])!;
                this._svgLine(svg, px, midY, cp.x + NODE_W / 2, cp.y);
            } else {
                // Multiple children: horizontal bar + individual drops
                const firstCp = positions.get(node.children[0])!;
                const lastCp  = positions.get(node.children[node.children.length - 1])!;
                const barX1   = firstCp.x + NODE_W / 2;
                const barX2   = lastCp.x  + NODE_W / 2;

                this._svgLine(svg, barX1, midY, barX2, midY);  // horizontal bar

                for (const child of node.children) {
                    const cp = positions.get(child)!;
                    const cx = cp.x + NODE_W / 2;
                    this._svgLine(svg, cx, midY, cx, cp.y);     // drop to each child
                }
            }
        }

        // ── Step 6: Render node boxes ─────────────────────────────────────────
        for (const [node, pos] of positions.entries()) {
            const box = document.createElement("div");
            box.className = "pos-node";
            if (node.isSelected) box.classList.add("pos-node--selected");
            if (node.isAncestor) box.classList.add("pos-node--ancestor");
            box.textContent    = node.name;
            box.title          = node.name;
            box.style.position  = "absolute";
            box.style.left      = `${pos.x}px`;
            box.style.top       = `${pos.y}px`;
            box.style.minWidth  = `${NODE_W}px`;
            box.style.minHeight = `${NODE_H}px`;
            box.style.boxSizing = "border-box";
            wrapper.appendChild(box);
        }

        this._chartContainer.appendChild(wrapper);
    }

    /** Helper: draw a coloured SVG line segment. */
    private _svgLine(
        svg: SVGSVGElement,
        x1: number, y1: number,
        x2: number, y2: number
    ): void {
        const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
        line.setAttribute("x1", x1.toString());
        line.setAttribute("y1", y1.toString());
        line.setAttribute("x2", x2.toString());
        line.setAttribute("y2", y2.toString());
        line.setAttribute("stroke", "#b3b0ad");
        line.setAttribute("stroke-width", "2");
        line.setAttribute("stroke-linecap", "round");
        svg.appendChild(line);
    }

    // ──────────────────────────────────────────────────────────────────────────
    // Status helpers
    // ──────────────────────────────────────────────────────────────────────────

    private _setStatus(message: string, type: "info" | "error" | "warning"): void {
        this._statusLabel.textContent = message;
        this._statusLabel.className   = `pos-chart-status pos-chart-status--${type}`;
    }
}
