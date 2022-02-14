###############################################################################
### Make DCD chart to html                                                  ###
### Module to install: pandas, plotly                                       ###
###############################################################################

import pandas as pd
import plotly.express as px

def result_to_chart(df_all, result_filename):

    # change phase name
    df_all.loc[df_all['Phase'] == 0, 'Phase'] = 'DIS0'
    df_all.loc[df_all['Phase'] == 2, 'Phase'] = 'CHARGE'
    df_all.loc[df_all['Phase'] == 4, 'Phase'] = 'DISCHARGE'

    # change time to H:MM:SS
    df_all['Time'] = pd.to_datetime(df_all['Time'], unit='s')#.dt.strftime('%H:%M:%S')

    # generate charts, seperate by Phase
    fig = px.line(df_all, x='Time', y='Voltage', color='Module', facet_col='Phase', title=result_filename+' DCD Chart')
    fig.update_xaxes(tickformat='%M:%S')

    # fig.show()

    # write fig to html
    fig.write_html(result_filename + '.html')
