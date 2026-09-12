"""Selected-cell reachability and reproducible shortest dependency paths."""
from collections import deque
from audit import REF, label, downstream
from openpyxl.utils.cell import range_boundaries


def select_cell(value, sheets):
    """Require one explicitly qualified A1 cell, never a range or external book."""
    match = REF.fullmatch(value)
    if not match:
        raise ValueError('Select one cell as Sheet!A1 or quoted sheet name followed by !A1')
    quoted, plain, cell = match.groups()
    name = quoted.replace("''", "'") if quoted is not None else plain
    if not name or any(c in name for c in '[]:') or ':' in cell:
        raise ValueError('Select one cell in this workbook, with its sheet name')
    names = {s['name'].casefold():s['name'] for s in sheets}
    if name.casefold() not in names: raise ValueError('Selected sheet does not exist')
    cell = cell.replace('$','').upper()
    col, row, _, _ = range_boundaries(cell)
    if col > 16384 or row > 1048576: raise ValueError('Selected cell is outside Excel bounds')
    return label(names[name.casefold()],cell)


def shortest_path(graph, source, target):
    """One shortest path; sorted neighbors break ties, independent of row order."""
    parents = {source:None}; queue = deque([source])
    while queue:
        node = queue.popleft()
        if node == target:
            path = []
            while node is not None:
                path.append(node); node = parents[node]
            return list(reversed(path))
        for child in sorted(graph.get(node, [])):
            if child not in parents:
                parents[child] = node; queue.append(child)
    return []


def selection(plan, source, target=None):
    source = select_cell(source,plan['sheets'])
    target = select_cell(target,plan['sheets']) if target is not None else None
    graph = plan['dependents']
    reachable = downstream(graph,source)
    path = shortest_path(graph,source,target) if target else []
    status = ('same_cell' if source==target else 'found' if path else 'no_parsed_path') if target else 'not_requested'
    return {'source':source,'target':target,'reachable_formulas':reachable,
            'path':path,'path_status':status,'edge_count':len(path)-1 if path else None,
            'unresolved_items_in_workbook':plan['coverage']['unresolved_items'],
            'interpretation':'One shortest path through parsed formula references, not all paths or proof of evaluated results. No parsed path does not prove independence when references are unsupported.'}
