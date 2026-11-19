"""Large vendor-style test model: a US national block plus 51 states, about 20,000 series.

    python examples/make_large_model.py large_model
    python model_map.py --eq large_model/equations.txt --a large_model/baseline.csv --b large_model/alt_scenario.csv -o large_map.html

National: prices (40 items), rates (20), output and employment by sector (60), income (40), housing (30).
Each state: population by age (inputs), employment and output by sector, income, housing, retail and 150
occupation series, each linked to its national counterpart and the state totals. Scenario B raises oil
35% and the Fed funds rate 0.5 from 2023Q1. Made-up numbers; standard library only.
"""
import random, sys, os
random.seed(7)
out = sys.argv[1] if len(sys.argv) > 1 else 'large_model'
os.makedirs(out, exist_ok=True)
ST = ['IAK','IAL','IAR','IAZ','ICA','ICO','ICT','IDC','IDE','IFL','IGA','IHI','IIA','IID','IIL','IIN','IKS','IKY','ILA','IMA',
      'IMD','IME','IMI','IMN','IMO','IMS','IMT','INC','IND','INE','INH','INJ','INM','INV','INY','IOH','IOK','IOR','IPA','IRI',
      'ISC','ISD','ITN','ITX','IUT','IVA','IVT','IWA','IWI','IWV','IWY']
eqs = []          # (name, const, [(driver, lag, coef)], rho)
exog = {}         # name -> base level
def E(n, c, drv, rho=0.0): eqs.append((n, c, drv, rho))
U = '.IUSA'
NS, NI, NP, NH, NR = 60, 40, 40, 30, 20          # sectors, income parts, price items, housing, rates
for x in ('FOIL', 'FRFED', 'FGOVX'): exog[x] = 1.0
# --- national block
E('FCPI'+U, 0.2, [('FOIL',0,0.15)] + [('FCPI%02d'%k+U,0,0.6/NP) for k in range(NP)], 0.0)
for k in range(NP): E('FCPI%02d'%k+U, 0.1, [('FOIL',0,0.05+0.1*random.random()), ('FCPI'+U,1,0.3)], 0.5)
E('FRGT10Y'+U, 0.1, [('FRFED',0,0.6), ('FCPI'+U,0,0.2)], 0.3)
for k in range(NR): E('FIR%02d'%k+U, 0.05, [('FRGT10Y'+U,0,0.5+0.3*random.random()), ('FRFED',0,0.2)], 0.2)
E('FGDP$'+U, 0.1, [('FOUT%02d'%k+U,0,0.8/NS) for k in range(NS)] + [('FGOVX',0,0.1)], 0.1)
for k in range(NS):
    E('FOUT%02d'%k+U, 0.1, [('FGDP$'+U,1,0.4), ('FIR%02d'%(k%NR)+U,0,-0.1), ('FCPI%02d'%(k%NP)+U,0,-0.05)], 0.4)
    E('FEMP%02d'%k+U, 0.1, [('FOUT%02d'%k+U,0,0.5)], 0.4)
E('FEMP'+U, 0.0, [('FEMP%02d'%k+U,0,1.0/NS) for k in range(NS)], 0.0)
for k in range(NI): E('FY%02d'%k+U, 0.1, [('FEMP'+U,0,0.5), ('FCPI'+U,0,0.2)], 0.3)
E('FYPDPI$'+U, 0.0, [('FY%02d'%k+U,0,1.0/NI) for k in range(NI)], 0.0)
for k in range(NH): E('FHS%02d'%k+U, 0.1, [('FYPDPI$'+U,0,0.4), ('FIR%02d'%(k%NR)+U,0,-0.2)], 0.4)
# --- states
for s in ST:
    S = '.'+s
    for k in range(20): exog['FPOP%02d'%k+S] = 1.0                       # population by age: inputs
    E('FPOP'+S, 0.0, [('FPOP%02d'%k+S,0,1/20) for k in range(20)], 0.0)
    for k in range(NS):
        E('FEMP%02d'%k+S, 0.05, [('FEMP%02d'%k+U,0,0.35), ('FEMP'+S,1,0.1), ('FPOP'+S,0,0.05)], 0.5)
        E('FOUT%02d'%k+S, 0.05, [('FEMP%02d'%k+S,0,0.4), ('FOUT%02d'%k+U,0,0.2)], 0.4)
    E('FEMP'+S, 0.0, [('FEMP%02d'%k+S,0,1/NS) for k in range(NS)], 0.0)
    E('FGSP'+S, 0.0, [('FOUT%02d'%k+S,0,1/NS) for k in range(NS)], 0.0)
    for k in range(NI): E('FY%02d'%k+S, 0.05, [('FEMP'+S,0,0.3), ('FY%02d'%k+U,0,0.2)], 0.4)
    E('FYPI'+S, 0.0, [('FY%02d'%k+S,0,1/NI) for k in range(NI)], 0.0)
    for k in range(NH): E('FHS%02d'%k+S, 0.05, [('FYPI'+S,0,0.3), ('FHS%02d'%k+U,0,0.2), ('FPOP'+S,0,0.1)], 0.4)
    E('FHPI'+S, 0.0, [('FHS%02d'%k+S,0,1/NH) for k in range(NH)], 0.0)
    for k in range(20): E('FRET%02d'%k+S, 0.05, [('FYPI'+S,0,0.4), ('FCPI%02d'%k+U,0,-0.1)], 0.3)
    for k in range(150): E('FOCC%03d'%k+S, 0.02, [('FEMP%02d'%(k%NS)+S,0,0.3), ('FYPI'+S,0,0.05)], 0.5)
names = [e[0] for e in eqs]
print('equations', len(eqs), 'inputs', len(exog), 'series', len(eqs)+len(exog))
T = 40
per = ['%dQ%d' % (2016 + t//4, t%4 + 1) for t in range(T)]
def simulate(shock):
    val = {}
    for x, b in exog.items():
        val[x] = [b * (1 + 0.004*t) for t in range(T)]
    if shock:
        val['FOIL'] = [v * (1.35 if t >= 28 else 1) for t, v in enumerate(val['FOIL'])]
        val['FRFED'] = [v + (0.5 if t >= 28 else 0) for t, v in enumerate(val['FRFED'])]
    for n in names: val[n] = [1.0]*T
    for t in range(T):
        for sweep in range(4):
            for n, c, drv, rho in eqs:
                y = c + rho * (val[n][t-1] if t else 1.0)
                for d, lag, co in drv:
                    y += co * val[d][t-lag if t-lag >= 0 else 0]
                val[n][t] = y
    return val
for fn, sh in (('baseline.csv', False), ('alt_scenario.csv', True)):
    v = simulate(sh)
    with open(os.path.join(out, fn), 'w') as f:
        f.write('name,' + ','.join(per) + '\n')
        for n in list(exog) + names:
            f.write(n + ',' + ','.join('%.6g' % x for x in v[n]) + '\n')
def term(d, lag, co):
    return '%+g*%s%s' % (co, d, '(-%d)' % lag if lag else '')
with open(os.path.join(out, 'equations.txt'), 'w') as f:
    for n, c, drv, rho in eqs:
        rhs = '%g' % c + ''.join(' ' + term(d, l, co) for d, l, co in drv) + (' %+g*%s(-1)' % (rho, n) if rho else '')
        f.write('%s = %s\n' % (n, rhs.replace('+-', '-')))
