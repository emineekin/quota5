import pandas as pd
from bokeh.plotting import figure, curdoc
from bokeh.layouts import widgetbox
from bokeh.models.widgets import Select
from bokeh.layouts import column, row
from bokeh.models import ColumnDataSource, NumeralTickFormatter, Range1d
from bokeh.models.tools import HoverTool
from pandas import ExcelWriter
from pandas import ExcelFile
import copy


class Data:
    def __init__(self):
        self.scholarship_levels = [100, 50, 0]
        self.df = pd.read_excel('ışık_yks_2020.xlsx', sheet_name='Yerleşenler')
        self.programs = self.df['PROGRAM'].unique()

    def get_students(self, program):
        ranks_for_program = {}
        x = 0
        total = len(self.df.loc[(self.df['PROGRAM'] == program)].index)
        for scholarship in self.scholarship_levels:
            students = self.df.loc[(self.df['PROGRAM'] == program) &
                                   (self.df['İNDİRİM'] == scholarship)]
            sorted_students = students.sort_values('SIRA', ascending=True)
            entry_ranks = [i for i in range(x + 1, x + len(sorted_students.index) + 1)]
            x += len(sorted_students.index)
            prefs = students['TERCİH']
            ranks = students['SIRA']
            ranks_for_program[scholarship] = {
                'x': entry_ranks,
                'ranks': ranks.tolist(),
                'pref': prefs.tolist(),
                'group': [len(entry_ranks) for i in entry_ranks],
                'total': [total for i in entry_ranks]
            }

        return ranks_for_program


class QuotaInterface:
    def __init__(self, in_data):
        self.data = data
        self.source_dict = {
            'x': [],
            'ranks': [],
            'pref': [],
            'group': [],
            'total': [],
        }
        self.plot_width = 800
        self.plot_height_rank = 400
        self.plot_height_pref = 200
        TOOLTIPS_QUOTA = [
            ("Kaçıncı", "@x"),
            ("Sıralama", "@ranks{0,0}"),
            ("Grup", "@group"),
            ("Toplam", "@total")
        ]
        hover_rank = HoverTool(tooltips=TOOLTIPS_QUOTA,
                               mode='vline')
        plot_options_rank = dict(plot_height=self.plot_height_rank,
                                 plot_width=self.plot_width,
                                 tools=[hover_rank],
                                 sizing_mode="fixed",
                                 toolbar_location=None,
                                 title='Sıralama',
                                 x_axis_label='Kaçıncı öğrencimiz',
                                 x_range=(0, 100),
                                 y_range=(0, 100)

                                 )
        self.figure_rank = figure(**plot_options_rank)
        self.figure_rank.yaxis.formatter = NumeralTickFormatter(format='0,0')

        TOOLTIPS_PREF = [
            ("Kaçıncı", "@x"),
            ("Tercih sırası", "@pref")
        ]
        hover_pref = HoverTool(tooltips=TOOLTIPS_PREF,
                               mode='vline')
        plot_options_pref = dict(plot_height=self.plot_height_pref,
                                 plot_width=self.plot_width,
                                 tools=[hover_pref],
                                 sizing_mode="fixed",
                                 toolbar_location=None,
                                 title='Tercih sırası',
                                 x_range=(0, 100),
                                 y_range=(0, 100)
                                 )
        self.figure_pref = figure(**plot_options_pref)
        self.figure_pref.y_range = Range1d(0, 25)

        self.lines_rank = {}
        self.lines_pref = {}

        self.colors = {
            100: 'green', 50: 'orange', 0: 'red'}

    def make_layout(self):
        select = Select(title='PROGRAM',
                        value=self.data.programs.tolist()[0],
                        options=self.data.programs.tolist())
        select.on_change('value', update_plot)

        for scholarship in self.data.scholarship_levels:
            source = ColumnDataSource(data=copy.deepcopy(self.source_dict))
            line_rank = self.figure_rank.line(x='x', y='ranks',
                                              line_width=3,
                                              color=self.colors[scholarship],
                                              source=source,
                                              legend_label='%d' % scholarship)
            line_pref = self.figure_pref.circle(x='x', y='pref',
                                                size=10,
                                                color=self.colors[scholarship],
                                                source=source)
            self.lines_rank[scholarship] = line_rank
            self.lines_pref[scholarship] = line_pref

        self.figure_rank.ray(x=0.1,
                             y=300000,
                             length=0,
                             angle=0,
                             angle_units='deg',
                             color='black',
                             line_width=2,
                             line_dash=[3, 2],
                             legend_label='300,000 sınırı')

        figures = column(self.figure_rank, self.figure_pref)
        layout_all = column(select, figures)
        curdoc().add_root(layout_all)

    def update(self, program):
        self.figure_rank.legend.location = 'bottom_right'
        max_y = 0
        max_x = 0
        ranks = self.data.get_students(program)
        for scholarship in self.data.scholarship_levels:
            self.lines_rank[scholarship].data_source.data = ranks[scholarship]
            self.lines_pref[scholarship].data_source.data = ranks[scholarship]
            if len(ranks[scholarship]['x']) > 0:
                max_x = max(max_x, max(ranks[scholarship]['x']))
            if len(ranks[scholarship]['ranks']) > 0:
                max_y = max(max_y, max(ranks[scholarship]['ranks']))
        extend = 1.4

        self.figure_rank.x_range.start = 0
        self.figure_rank.x_range.end = max_x * extend

        self.figure_rank.y_range.start = 0
        self.figure_rank.y_range.end = max_y * extend

        self.figure_pref.x_range.start = 0
        self.figure_pref.x_range.end = max_x * extend


def update_plot(attr, old, new):
    interface.update(new)


data = Data()
interface = QuotaInterface(data)
interface.make_layout()
update_plot(None, None, data.programs.tolist()[0])

