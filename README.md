# model_map

Group the variables of an equation model and draw a map you can unfold, like zooming a map.

Give it the model's equations (EViews style, `Y = formula`). It finds groups of variables that are densely linked inside and loosely linked to the rest. Each group is split again until every view fits on one screen. It writes one HTML page with:

- **All variables, unfold step by step.** One map of the whole model that starts folded. Click a group to unfold it in place: its sub-groups appear in the same map, linked to everything else, and you can go down to single variables while the other groups stay folded.
- **Main chain, zoomable.** The route through the model, like directions on a map. Click a step to see the chain inside it.
- **Find a variable.** Unfolds the map down to it and shows its equation, its inputs, and what uses it.
- **All variables, by group.** A table of every variable, folded by the same groups.

Scenario data is optional. With `--a` and `--b`, boxes are coloured by the gap B − A.

Standard-library Python only (3.8+). No pip install. Reads .txt/.prg equation files and .csv/.xlsx data.

## Quick start

```bash
python model_map.py --eq model.txt -o map.html                                   # structure only
python model_map.py --eq model.txt --a base.csv --b scenario.xlsx -o map.html    # coloured by the gap
python model_map.py --eq model.txt --groups --depth 2                            # also print the group tree
```

Open `map.html` in any browser.

Examples:

```bash
python model_map.py --eq examples/macro/equations.txt --a examples/macro/baseline.csv --b examples/macro/alt_scenario.xlsx -o examples/macro_map.html
python model_map.py --eq examples/bank/equations_base.txt --eq-extra examples/bank/equations_stress.txt -o examples/bank_map.html
python examples/make_large_model.py large_model      # a 20,000-series test model, US + 51 states
python model_map.py --eq large_model/equations.txt --a large_model/baseline.csv --b large_model/alt_scenario.csv -o large_map.html
```

On the 20,000-series model the map is built in about 6 seconds as a 3.5 MB page. The page opens in under half a second, and each click takes about 0.2 seconds.

## How the groups are found

Nothing is hard-coded: the groups come from the equation links alone.

1. **Links.** If X appears in Y's equation (at any lag), X and Y are linked.
2. **Groups (Louvain community detection).** The split maximises modularity:

   Q = (1/2m) · Σᵢⱼ [ Aᵢⱼ − γ · kᵢkⱼ/2m ] · 1(i, j in the same group)

   - Aᵢⱼ = 1 if i and j are linked.
   - kᵢ = the number of links of i; m = the total number of links.

   Q is high when groups have more links inside than chance would give. The resolution γ is tuned for each split to prefer balanced parts.
3. **Levels.** Each group is split again until no view has more than `--map-boxes` boxes (default 40).
   - Too many parts: groups that hang off the same neighbours are put together.
   - No structure at all: the set is cut along the order of its equations.
4. **Key variables.** For each group, the fewest variables that every link between it and the rest of the model passes through: a minimum vertex cover of its boundary links (König's theorem). A group joined through one or two key variables is almost a separate model hanging off them.
5. **Names.** Geography codes found in the names (e.g. `ICA`, `_ICAN`) are used only to name groups, such as "ICA" or "IKY + IID + IUSA". Otherwise a group is named after its most-connected variable ("around FGDP$.IUSA"). `--map-geo` splits by geography first instead.

Variables that appear in no equation and have none of their own go in a "(not linked)" box.

## The main chain

Each link u → v gets a weight, the search path count:

  SPC(u→v) = N⁻(u) · N⁺(v)

N⁻(u) is the number of distinct chains from any input (a variable with no equation) to u, and N⁺(v) the number from v to any end variable. So SPC counts the input-to-end chains that run through the link (main path analysis, Hummon & Doreian 1989).

- **At the top:** the chain starts at the heaviest link between groups and extends along the heaviest links. It stops when the next link weighs less than 5% of the first.
- **Inside a step:** the chain starts where links from the step before enter the group and ends where links leave for the next step. It's the widest path, the route whose weakest link is strongest. So each level continues the same route.

**Layers** = the longest chain of equations downstream of a variable. A loop counts as one layer.

## Options

| Option | Default | |
|---|---|---|
| `--eq FILE` | | equation file(s) or folder; repeatable |
| `--eq-extra FILE` | | more equations that override `--eq`, e.g. a second version of the model |
| `--a FILE`, `--b FILE` | | optional scenario data, to colour boxes by the gap B − A |
| `--exog NAMES` | | variables to treat as exogenous (names, or a file of names) |
| `--map-boxes N` | 40 | most boxes in one view; groups get more levels to fit |
| `--map-geo` | off | split by geography first instead of by links |
| `--geo-sep` | auto | separator before geography codes in names: `.`, `_` or auto |
| `--no-geo-names` | off | don't use geographies to name groups |
| `--groups`, `--depth N` | | also print the group tree in the terminal |
| `--title`, `-o FILE`, `-q` | | page title, output file, quiet |

## Equations

The same reader as scenario_compare:
- `Y = formula`, `log(Y) = …`, `d(Y) = …` and other EViews forms;
- comments (`'`);
- error terms `[ar(1)=…, …]`;
- lines that continue on the next line;
- scientific notation such as `1.5E -05`;
- equations in spreadsheets.

A line that cannot be read is counted on the page, and its variable keeps the links found in its text.
