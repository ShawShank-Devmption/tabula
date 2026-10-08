"""Dependency graph over formula cells (design.md section 6.2).

Keys are (sheet_id, col, row).  Single-cell precedents are stored as edges;
range precedents are stored symbolically as bounds, so a whole-column range
costs one entry instead of 1,048,576.  Every algorithm here is iterative:
real workbooks have dependency chains far deeper than Python's recursion limit.
"""
from __future__ import annotations

from collections import defaultdict, deque

Key = tuple  # (sid, col, row)
Bounds = tuple  # (c1, r1, c2, r2)
WIDE = 64  # ranges spanning more columns than this live in a per-sheet list instead of buckets


def _inside(col: int, row: int, b: Bounds) -> bool:
    return b[0] <= col <= b[2] and b[1] <= row <= b[3]


class DependencyGraph:
    def __init__(self):
        self.cell_dependents: dict[Key, set[Key]] = defaultdict(set)
        # range edges, indexed by column so a lookup only scans ranges over that column:
        # sid -> col -> [(r1, r2, dependent)];  very wide ranges: sid -> [(bounds, dependent)]
        self.range_index: dict[int, dict[int, list]] = defaultdict(lambda: defaultdict(list))
        self.wide_ranges: dict[int, list] = defaultdict(list)
        self.precedents: dict[Key, tuple[set[Key], list[tuple[int, Bounds]]]] = {}
        self.formulas_by_sheet: dict[int, set[tuple[int, int]]] = defaultdict(set)

    # -- maintenance ---------------------------------------------------------
    def set_formula(self, key: Key, cells: set[Key], ranges: list[tuple[int, Bounds]]) -> None:
        self.remove_formula(key)
        self.precedents[key] = (set(cells), list(ranges))
        self.formulas_by_sheet[key[0]].add((key[1], key[2]))
        for c in cells:
            self.cell_dependents[c].add(key)
        for sid, b in ranges:
            if b[2] - b[0] < WIDE:
                for col in range(b[0], b[2] + 1):
                    self.range_index[sid][col].append((b[1], b[3], key))
            else:
                self.wide_ranges[sid].append((b, key))

    def remove_formula(self, key: Key) -> None:
        old = self.precedents.pop(key, None)
        if old is None:
            return
        self.formulas_by_sheet[key[0]].discard((key[1], key[2]))
        cells, ranges = old
        for c in cells:
            deps = self.cell_dependents.get(c)
            if deps:
                deps.discard(key)
        for sid, b in ranges:
            if b[2] - b[0] < WIDE:
                for col in range(b[0], b[2] + 1):
                    bucket = self.range_index[sid][col]
                    bucket[:] = [e for e in bucket if e[2] != key]
            else:
                self.wide_ranges[sid] = [e for e in self.wide_ranges[sid] if e[1] != key]

    def is_formula(self, key: Key) -> bool:
        return key in self.precedents

    # -- queries -------------------------------------------------------------
    def dependents(self, key: Key) -> set[Key]:
        """Formulas that read `key` directly (cell edges + containing ranges)."""
        out = set(self.cell_dependents.get(key, ()))
        sid, col, row = key
        index = self.range_index.get(sid)
        if index is not None:
            for r1, r2, dep in index.get(col, ()):
                if r1 <= row <= r2:
                    out.add(dep)
        for b, dep in self.wide_ranges.get(sid, ()):
            if _inside(col, row, b):
                out.add(dep)
        return out

    def formulas_in(self, sid: int, b: Bounds) -> list[Key]:
        fset = self.formulas_by_sheet.get(sid, set())
        size = (b[2] - b[0] + 1) * (b[3] - b[1] + 1)
        if size <= len(fset):
            return [(sid, c, r) for r in range(b[1], b[3] + 1)
                    for c in range(b[0], b[2] + 1) if (c, r) in fset]
        return [(sid, c, r) for (c, r) in fset if _inside(c, r, b)]

    def formula_precedents(self, key: Key) -> set[Key]:
        cells, ranges = self.precedents.get(key, ((), ()))
        out = {c for c in cells if c in self.precedents}
        for sid, b in ranges:
            out.update(self.formulas_in(sid, b))
        return out

    def closure(self, seeds) -> set[Key]:
        """Seeds plus every formula transitively depending on them (BFS)."""
        seen: set[Key] = set()
        queue = deque(seeds)
        while queue:
            k = queue.popleft()
            if k in seen:
                continue
            seen.add(k)
            queue.extend(d for d in self.dependents(k) if d not in seen)
        return {k for k in seen if k in self.precedents}

    # -- scheduling ----------------------------------------------------------
    def order(self, nodes: set[Key]) -> tuple[list[Key], set[Key]]:
        """Topological order of `nodes` (Kahn) and the set of cycle members (Tarjan)."""
        preds = {n: self.formula_precedents(n) & nodes for n in nodes}
        order = _kahn(nodes, preds)
        cyclic: set[Key] = set()
        remaining = nodes - set(order)
        if remaining:
            for scc in _tarjan(remaining, preds):
                if len(scc) > 1 or scc[0] in preds[scc[0]]:
                    cyclic.update(scc)
            rest = remaining - cyclic
            order.extend(_kahn(rest, {n: preds[n] & rest for n in rest}))
        return order, cyclic

    def cycle_path(self, start: Key) -> list[Key] | None:
        """A path start -> ... -> start following 'reads' edges, or None."""
        parent: dict[Key, Key] = {}
        stack = [start]
        visited = {start}
        while stack:
            node = stack.pop()
            for p in sorted(self.formula_precedents(node)):
                if p == start:
                    return _path(start, node, parent)
                if p not in visited:
                    visited.add(p)
                    parent[p] = node
                    stack.append(p)
        return None


def _path(start: Key, last: Key, parent: dict[Key, Key]) -> list[Key]:
    chain = [last]
    while chain[-1] != start:
        chain.append(parent[chain[-1]])
    chain.reverse()          # start, ..., last
    return chain + [start]   # last reads start: the cycle closes


def _kahn(nodes, preds) -> list[Key]:
    indeg = {n: len(preds[n]) for n in nodes}
    succ: dict[Key, list[Key]] = defaultdict(list)
    for n in nodes:
        for p in preds[n]:
            succ[p].append(n)
    queue = deque(sorted(n for n in nodes if indeg[n] == 0))
    order = []
    while queue:
        n = queue.popleft()
        order.append(n)
        for s in sorted(succ[n]):
            indeg[s] -= 1
            if indeg[s] == 0:
                queue.append(s)
    return order


def _tarjan(nodes, preds) -> list[list[Key]]:
    """Strongly connected components, iteratively (no recursion)."""
    index: dict[Key, int] = {}
    low: dict[Key, int] = {}
    on_stack: set[Key] = set()
    stack: list[Key] = []
    sccs: list[list[Key]] = []
    counter = 0
    for root in sorted(nodes):
        if root in index:
            continue
        work = [(root, iter(sorted(preds[root] & nodes)))]
        index[root] = low[root] = counter
        counter += 1
        stack.append(root)
        on_stack.add(root)
        while work:
            node, children = work[-1]
            advanced = False
            for child in children:
                if child not in index:
                    index[child] = low[child] = counter
                    counter += 1
                    stack.append(child)
                    on_stack.add(child)
                    work.append((child, iter(sorted(preds[child] & nodes))))
                    advanced = True
                    break
                if child in on_stack:
                    low[node] = min(low[node], index[child])
            if advanced:
                continue
            work.pop()
            if work:
                parent = work[-1][0]
                low[parent] = min(low[parent], low[node])
            if low[node] == index[node]:
                scc = []
                while True:
                    w = stack.pop()
                    on_stack.discard(w)
                    scc.append(w)
                    if w == node:
                        break
                sccs.append(sorted(scc))
    return sccs
