# model_map

Group the variables of an equation model and draw a map you can unfold, like zooming a map.

Give it the model's equations (EViews style, `Y = formula`), or just a list of which variable feeds which. It finds groups of variables that are densely linked inside and loosely linked to the rest. Each group is split again until every view fits on one screen. It writes one HTML page with four views:

- **All variables, unfold step by step.** One map of the whole model that starts folded. Click a group to unfold it in place: its sub-groups appear in the same map, linked to everything else, and you can go down to single variables while the other groups stay folded.
- **Main chain, zoomable.** The route through the model, like directions on a map. Click a step to see the chain inside it.
- **Around a variable.** What feeds a variable and what it feeds, 0–4 equations away in each direction. Double-click a box to put it in the middle.
- **All variables, by group.** A table of every variable, folded by the same groups.

**Find a variable** by name or description: a list of matches appears as you type. Picking one unfolds the map down to it and opens its panel, which shows the equation, what it uses, what uses it, its groups, layers and gap.

**Optional extras:**
- Scenario data (`--a`, `--b`) colours boxes by the gap B − A. Descriptions in the data files are picked up automatically; `--desc` adds more.

Standard-library Python only (3.8+). No pip install. It's one file, so to use it in another project, copy `model_map.py` there and `import model_map`.

## Quick start

```bash
python model_map.py --eq model.txt -o map.html                                   # structure only
python model_map.py --eq model.txt --a base.csv --b scenario.xlsx -o map.html    # coloured by the gap
python model_map.py --links links.csv -o map.html                                # any model: a table of links
python model_map.py --eq model.txt --json map.json --csv variables.csv           # data for other tools
python model_map.py --eq model.txt --groups --depth 2                            # print the group tree
```

Open `map.html` in any browser.

## Use it from Python (other projects)

```python
import model_map as mm                         # model_map.py next to your script

m = mm.ModelMap.from_equations(['model.txt'])               # EViews equations
m = mm.ModelMap.from_links([('FOIL', 'FCPI'), ('FCPI', 'FGDP')])   # (from, to) pairs from anywhere
m = mm.ModelMap.from_links_csv('links.csv')                 # columns from/to, source/target or driver/variable
m = mm.ModelMap({'FGDP': ['FC', 'FI'], 'FC': ['FYPDPI']})   # variable -> what its equation uses

m.add_scenarios(['base.csv'], ['scen.xlsx'])   # optional: gaps B - A for colours
m.set_gaps({'FGDP': -1.4})                     # ... or your own numbers
m.set_descriptions('descriptions.csv')         # ... and labels

m.groups()                      # [{'name', 'size', 'keys', 'members', 'groups': [...]}, ...]
m.group_path('FHPI.ICA')        # ['around FYPDPI$.IWY', 'ICA + ICO']
m.key_variables('ICA + ICO')    # the fewest variables its outside links pass through
m.upstream('FHPI.ICA', 2)       # {'FHPI.IUSA': 1, 'FYPDPI$.ICA': 1, ...}   what feeds it, by steps
m.downstream('FOIL')            # what it feeds
m.link_path('FOIL', 'FHPI.ICA') # ['FOIL', 'FCPI.IUSA', 'FRGT10Y.IUSA', 'FHPI.IUSA', 'FHPI.ICA']
m.layers()['FOIL']              # longest chain of equations downstream
m.table()                       # one row per variable with its group path
m.summary()

m.to_html('map.html'); m.to_json('map.json'); m.to_csv('variables.csv')
```

`examples/use_from_python.py` runs all of this on the example model. Options for the groups go to the constructor: `boxes=40`, `by_geo=False`, `geo_sep='auto'`, `geo_names=True`, `exog=[...]`.

To install it instead (where your environment allows): `pip install git+https://github.com/wuxiaokai-hk/model_map`. That gives `import model_map` and a `model-map` command.

## Examples

```bash
python model_map.py --eq examples/macro/equations.txt --a examples/macro/baseline.csv --b examples/macro/alt_scenario.xlsx -o examples/macro_map.html
python model_map.py --eq examples/bank/equations_base.txt --eq-extra examples/bank/equations_stress.txt -o examples/bank_map.html
python examples/make_large_model.py large_model      # a 20,000-series test model, US + 51 states
python model_map.py --eq large_model/equations.txt --a large_model/baseline.csv --b large_model/alt_scenario.csv -o large_map.html
```

On the 20,000-series model the map is built in about 5 seconds as a 3.5 MB page. The page opens in under half a second, and each click takes about 0.2 seconds.

## A variable is missing from the map?

```bash
python model_map.py --eq model.txt --a base.xlsx --b scen.xlsx --find FGDP$.ICAN
```

This shows where the name appears and what happened to it:
- the equation lines that mention it;
- each data file, sheet, row and column where it appears, and how many numbers that row has;
- whether it was loaded, with the reader's notes for each sheet;
- if it's nowhere, the closest names (e.g. a typo, or a different separator).

`--unread` lists equation lines that could not be read.

Every sheet of an Excel file is read; sheets without dates (notes, metadata) are skipped and listed. A series with no numbers is still shown in the map, just without a gap.

## Tests

```bash
python -m unittest discover -s tests
```

The tests check that:
- every variable sits in exactly one group, and sub-groups split their parent exactly;
- every view fits the box limit;
- every link between a group and the rest touches one of its key variables;
- loops count as one layer;
- the page, JSON and CSV outputs are written correctly.

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
5. **Names.** Each group is named by what is in it.
   - **Concepts:** taken from the names, with the geography removed and sector numbers dropped (FEMP29.ICA → FEMP). When the data files have descriptions, those are used instead.
   - **Regions:** e.g. "FOCC, FEMP, FOUT · ICA", "House price index · 13 regions", "FY · IUSA + 50 regions".
   - **Mixed groups:** e.g. "6 concepts: Employment, House price index, … · 42 regions".

   Geographies are found in the names automatically and are used only for naming. `--map-geo` splits by geography first instead.
6. **Add-factor series.** Series like `FGDP.IUSA_A` or `FGDP_A.IUSA` that no equation uses are left out of the map and the table. Each is listed in its variable's panel. The suffix is set with `--af-suffix` (default `_A`; `''` keeps them).

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
| `--links FILE` | | instead of equations: a csv/xlsx of links (from/to columns); repeatable |
| `--desc FILE` | | csv/xlsx of variable descriptions (name, description) |
| `--json FILE`, `--csv FILE` | | also write groups, layers, links and gaps as JSON; every variable with its group path as CSV |
| `--eq-extra FILE` | | more equations that override `--eq`, e.g. a second version of the model |
| `--a FILE`, `--b FILE` | | optional scenario data, to colour boxes by the gap B − A |
| `--exog NAMES` | | variables to treat as exogenous (names, or a file of names) |
| `--map-boxes N` | 40 | most boxes in one view; groups get more levels to fit |
| `--af-suffix` | `_A` | add-factor series (`X_A`, `STEM_A.GEO`) that no equation uses are shown with their variable, not as boxes; `''` keeps them |
| `--map-geo` | off | split by geography first instead of by links |
| `--geo-sep` | auto | separator before geography codes in names: `.`, `_` or auto |
| `--no-geo-names` | off | don't use geographies to name groups |
| `--groups`, `--depth N` | | also print the group tree in the terminal |
| `--title`, `-o FILE`, `-q` | | page title, output page (default `model_map.html`), quiet |

## Equations

The same reader as scenario_compare:
- `Y = formula`, `log(Y) = …`, `d(Y) = …` and other EViews forms;
- comments (`'`);
- error terms `[ar(1)=…, …]`;
- lines that continue on the next line;
- scientific notation such as `1.5E -05`;
- equations in spreadsheets.

A line that cannot be read is counted on the page, and its variable keeps the links found in its text.

## License

MIT, see [LICENSE](LICENSE). Free to use, change and share. Data downloaded from public sources stay under their providers' terms. Independent open-source project, not affiliated with or endorsed by any employer, regulator or data vendor.
