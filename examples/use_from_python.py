"""Using model_map from another script or project.

Copy model_map.py next to your script (or add its folder to sys.path), then import it.
Run from the repo root:  python examples/use_from_python.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import model_map as mm

HERE = os.path.dirname(os.path.abspath(__file__))

# 1. From equation files (EViews style), with scenario data to colour the map
m = mm.ModelMap.from_equations([os.path.join(HERE, 'macro', 'equations.txt')])
m.add_scenarios([os.path.join(HERE, 'macro', 'baseline.csv')], [os.path.join(HERE, 'macro', 'alt_scenario.xlsx')])
print(m.summary())

# 2. Questions about the model
print('Groups of FHPI.ICA:', ' > '.join(m.group_path('FHPI.ICA')))
print('What feeds it within 2 steps:', m.upstream('FHPI.ICA', 2))
print('How oil reaches it:', ' -> '.join(m.link_path('FOIL', 'FHPI.ICA')))
print('Layers of FOIL:', m.layers()['FOIL'])
top = m.groups()[0]
print('Biggest group:', top['name'], top['size'], 'variables, key variables:', top['keys'])

# 3. Any other model: give the links yourself (from, to) or a dict variable -> what it uses
links = [('GDPG', 'LOANS'), ('LOANS', 'RWA'), ('RWA', 'CET1'), ('UNEMP', 'PD'), ('PD', 'ECL'), ('ECL', 'CET1')]
small = mm.ModelMap.from_links(links)
small.set_gaps({'LOANS': -3.1, 'RWA': 2.0, 'CET1': -0.8})       # your own numbers to colour the boxes
small.set_descriptions({'CET1': 'Common equity tier 1 ratio'})
print('From UNEMP to CET1:', small.link_path('UNEMP', 'CET1'))

# 4. Outputs
m.to_html('macro_map.html')        # the interactive page
m.to_json('macro_map.json')        # groups, layers, links, gaps
m.to_csv('macro_variables.csv')    # every variable with its group path
small.to_html('small_map.html', title='Capital model')
print('Written: macro_map.html, macro_map.json, macro_variables.csv, small_map.html')
