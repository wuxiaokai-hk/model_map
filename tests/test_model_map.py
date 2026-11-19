"""Tests for model_map. Run: python -m unittest discover -s tests   (standard library only)."""
import json
import os
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)
import model_map as mm  # noqa: E402

EX = os.path.join(ROOT, 'examples')


def chain_model():
    # X -> A -> B -> C, plus a loop B <-> D, and Z alone feeding C
    return mm.ModelMap({'A': ['X'], 'B': ['A', 'D'], 'D': ['B'], 'C': ['B', 'Z']})


class TestGraph(unittest.TestCase):
    def test_links_and_walks(self):
        m = chain_model()
        self.assertEqual(m.parents('B'), ['A', 'D'])
        self.assertEqual(sorted(m.children('B')), ['C', 'D'])
        self.assertEqual(m.upstream('C', 1), {'B': 1, 'Z': 1})
        self.assertEqual(set(m.upstream('C')), {'A', 'B', 'D', 'X', 'Z'})
        self.assertEqual(m.downstream('X'), {'A': 1, 'B': 2, 'C': 3, 'D': 3})
        self.assertEqual(m.link_path('X', 'C'), ['X', 'A', 'B', 'C'])
        self.assertIsNone(m.link_path('C', 'X'))

    def test_layers_loop_counts_once(self):
        L = chain_model().layers()
        # X -> A -> {B,D loop} -> C : 3 layers from X, the loop counted once
        self.assertEqual(L['X'], 3)
        self.assertEqual(L['B'], L['D'])
        self.assertEqual(L['C'], 0)

    def test_from_links_matches_dict(self):
        a = mm.ModelMap.from_links([('X', 'A'), ('A', 'B'), ('D', 'B'), ('B', 'D'), ('B', 'C'), ('Z', 'C')])
        self.assertEqual(a.layers(), chain_model().layers())

    def test_links_csv(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, 'links.csv')
            with open(p, 'w') as f:
                f.write('source,target\nX,A\nA,B\nB,C\n')
            m = mm.ModelMap.from_links_csv(p)
            self.assertEqual(m.link_path('X', 'C'), ['X', 'A', 'B', 'C'])

    def test_exog_cuts_its_equation(self):
        m = mm.ModelMap({'A': ['X'], 'B': ['A']}, exog=['A'])
        self.assertEqual(m.parents('A'), [])
        self.assertEqual(m.upstream('B'), {'A': 1})


class TestGroups(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.m = mm.ModelMap.from_equations([os.path.join(EX, 'macro', 'equations.txt')], boxes=20)

    def all_groups(self, gs=None):
        for g in (self.m.groups() if gs is None else gs):
            yield g
            yield from self.all_groups(g.get('groups') or [])

    def test_every_variable_in_exactly_one_top_group(self):
        top = self.m.groups()
        seen = [n for g in top for n in g['members']]
        self.assertEqual(sorted(seen), self.m.nodes)
        self.assertEqual(len(seen), len(set(seen)))

    def test_children_split_their_parent(self):
        for g in self.all_groups():
            if g.get('groups'):
                kids = sorted(n for c in g['groups'] for n in c['members'])
                self.assertEqual(kids, sorted(g['members']), g['name'])

    def test_views_fit(self):
        for g in self.all_groups():
            if g.get('groups'):
                self.assertLessEqual(len(g['groups']), 20, g['name'])
            else:
                self.assertTrue(len(g['members']) <= 20 or g['name'] == '(not linked)', g['name'])

    def test_keys_cover_every_boundary_link(self):
        # every link between a group and the rest touches one of its key variables
        for g in self.all_groups():
            if g['name'] == '(not linked)':
                continue
            mem, keys = set(g['members']), set(g['keys'])
            for v in mem:
                for u in self.m.parents(v):
                    if u not in mem:
                        self.assertTrue(u in keys or v in keys, (g['name'], u, v))

    def test_group_path_and_table(self):
        path = self.m.group_path('FHPI.ICA')
        self.assertTrue(path)
        rows = self.m.table()
        self.assertEqual(len(rows), len(self.m.nodes))


class TestOutputs(unittest.TestCase):
    def test_html_json_csv(self):
        m = mm.ModelMap.from_equations([os.path.join(EX, 'bank', 'equations_base.txt')],
                                       extra=[os.path.join(EX, 'bank', 'equations_stress.txt')])
        m.add_scenarios([os.path.join(EX, 'bank', 'base.csv')], [os.path.join(EX, 'bank', 'stress.csv')])
        m.set_descriptions({'PD': 'Probability of default'})
        with tempfile.TemporaryDirectory() as d:
            h, j, c = (os.path.join(d, x) for x in ('m.html', 'm.json', 'm.csv'))
            m.to_html(h)
            m.to_json(j)
            m.to_csv(c)
            html = open(h, encoding='utf-8').read()
            self.assertIn('Probability of default', html)
            self.assertNotIn('__DATA__', html)
            data = json.load(open(j))
            self.assertEqual(data['summary']['variables'], len(m.nodes))
            self.assertIn('PD', data['links']['ECL'])
            self.assertEqual(len(open(c).read().strip().splitlines()) - 1, len(m.nodes))
        self.assertTrue(m.gap)

    def test_command_line(self):
        with tempfile.TemporaryDirectory() as d:
            out = os.path.join(d, 'x.html')
            m = mm.build(mm_args(['--eq', os.path.join(EX, 'macro', 'equations.txt'), '-o', out, '-q']))
            self.assertTrue(os.path.getsize(out) > 10000)
            self.assertGreater(m.summary()['variables'], 300)


def mm_args(argv):
    import argparse
    ns = {}
    orig = mm.build
    try:
        mm.build = lambda a: ns.setdefault('a', a)
        mm.main(argv)
    finally:
        mm.build = orig
    return ns['a']


if __name__ == '__main__':
    unittest.main()
