import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
from ipywidgets import interact
import cufflinks as cf
import datetime as dt

cf.go_offline()

link1 = 'C:/Users/p119124/Documents/Historical_Propylene.xlsx'
tbl1 = pd.read_excel(link1, header=11, usecols="B:E")
tbl1.columns = ['Date', 'Propylene_Chemical_Low', 'Propylene_Chemical_High', 'Propylene_Chemical_Mid']
tbl2 = tbl1.loc[:1546,['Date','Propylene_Chemical_Mid']]

tbl2 = tbl2[tbl2['Date'] > dt.datetime.strptime('2006-01-01','%Y-%m-%d')]

tbl4 = tbl1.loc[:1546, ['Date', 'Propylene_Chemical_Mid']]
tbl4['std_up'] = tbl4.loc[:, 'Propylene_Chemical_Mid'] + tbl4.loc[:,'Propylene_Chemical_Mid'].rolling(50).std().fillna(0)
tbl4['std_up'] = tbl4['std_up'].rolling(window=10).mean()
tbl4['std_down'] = tbl4.loc[:,'Propylene_Chemical_Mid'] - tbl4.loc[:,'Propylene_Chemical_Mid'].rolling(50).std().fillna(0)
tbl4['std_down'] = tbl4['std_down'].rolling(window=10).mean()
tbl4 = pd.melt(tbl4, id_vars=["Date","Value_Type"], var_name="Std_Type", value_name="Value")
tbl4['Std_Type'] = tbl4['Std_Type'].replace('Propylene_Chemical_Mid','Mid')


x = tbl2['Propylene_Chemical_Mid'].median()
mean = tbl2['Propylene_Chemical_Mid'].mean()
std = tbl2['Propylene_Chemical_Mid'].std()

@interact(Width1=(0,1000))
def plot1(Width1):
    plt.figure(figsize=(12,6))
    sns.distplot(tbl2['Propylene_Chemical_Mid'], bins=50)
    plt.title('Propylene - Histogram')
    plt.vlines(x, 0, 0.0015, color='gray')
    plt.vlines(mean + std + Width1, 0, 0.0015, color='red')
    plt.vlines(mean - std, 0, 0.0015, color='red')
    plt.ylim([0, 0.0012])
    plt.plot([1000, 1000], [0.0002, 0.0008], 'k-', lw=2, linestyle='dashed')

class ChartsPropylene():
    def __init__(self):
        self.start_date=datetime(2008, 4, 24)
        self.end_date=datetime(2020, 5, 24)
        self.dates=pd.date_range(self.start_date, self.end_date, freq='D')
        self.options=[(date.strftime(' %d %b %y '), date) for date in self.dates]
        self.index=(0, len(options)-1)
        self.selection_range_slider=widgets.SelectionRangeSlider(options=self.options, index=self.index, description='Dates', orientation='horizontal', layout={'width':'600px'})

    def __printChart__(self):
        display(self.selection_range_slider)
        x=self.selection_range_slider.get_interact_value()[0].toordinal()
        y=self.selection_range_slider.get_interact_value()[1].toordinal()
        abs1=abs(y-x)
        plt.figure(figsize=(18, 10))
        sns.set(style="darkgrid")
        palette2=sns.color_palette("mako_r", 3)
        sns.lineplot(x="Date", y="Value", hue='Std_Type', style='Value_Type', sizes=(.25, 2.5), ci='sd', estimator=None, lw=1, palette=palette2, data=tbl4)
        rectangle1=plt.Rectangle(xy=(x, 500), width=abs1, height=500, linewidth=2, color='red', facecolor='blue', joinstyle='round', alpha=0.1, fill=True)
        rectangle2=plt.Rectangle(xy=(x, 500), width=abs1, height=500, linewidth=2, color='red', facecolor='blue', joinstyle='round', alpha=1, fill=False)
        plt.gca().add_patch(rectangle1)
        plt.gca().add_patch(rectangle2)
        plt.show()


ChartsPropylene().__printChart__()