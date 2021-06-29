import http.client
import json
import datetime
from openpyxl import Workbook
import plotly.offline
import numpy
import fpdf


class HTML2PDF(fpdf.FPDF, fpdf.HTMLMixin):
    pass


def team_function(team_id, connect, head):
    position = dict(Midfielder='Полузащитник',
                    Goalkeeper='Вратарь',
                    Defender='Защитник',
                    Attacker='Нападающий')
    team_players = ''
    connect.request('GET', '/v2/teams/' + team_id, None, head)
    response_teams = json.loads(connection.getresponse().read().decode())
    for j in range(len(response_teams['squad'])):
        if response_teams['squad'][j]['position'] is not None:
            team_players += position[response_teams['squad'][j]['position']] + ' ' + response_teams['squad'][j][
                'name'] + ', '
    if len(team_players) > 1:
        team_players = team_players[0:-2]
    return team_players


dateFrom = str(datetime.date.today() + datetime.timedelta(days=1))
dateTo = str(datetime.date.today() + datetime.timedelta(days=4))

connection = http.client.HTTPConnection('api.football-data.org')
headers = {'X-Auth-Token': '15095c41807a4cbb926b1a44cb4fd3ce'}
connection.request('GET', '/v2/competitions/EC/matches?dateFrom=' + dateFrom + '&dateTo=' + dateTo, None, headers)
response_matches = json.loads(connection.getresponse().read().decode())

matches = []

for i in range(len(response_matches['matches'])):
    if response_matches['matches'][i]['homeTeam']['name'] is not None:
        matches.append(dict(homeTeam=response_matches['matches'][i]['homeTeam']['name'],
                            homeTeam_id=response_matches['matches'][i]['homeTeam']['id'],
                            awayTeam=response_matches['matches'][i]['awayTeam']['name'],
                            awayTeam_id=response_matches['matches'][i]['awayTeam']['id'],
                            utcDate=response_matches['matches'][i]['utcDate'],
                            oddsHomeWin=response_matches['matches'][i]['odds']['homeWin'],
                            oddsDraw=response_matches['matches'][i]['odds']['draw'],
                            oddsAwayWin=response_matches['matches'][i]['odds']['awayWin']))

home_teams = []
away_teams = []

for i in range(len(matches)):
    if matches[i]['homeTeam'] is not None:
        home_teams.append(team_function(str(matches[i]['homeTeam_id']), connection, headers))
        away_teams.append(team_function(str(matches[i]['awayTeam_id']), connection, headers))

outputData = [['Home team', 'Guest team', 'Date and time(Moscow)', 'Odds home team',
               'Odds draw', 'Odds guest team']] # 'Состав домашней команды',
               # 'Состав гостевой команды']]

for i in range(len(matches)):
    outputData.append([matches[i]['homeTeam'], matches[i]['awayTeam'],
                       datetime.datetime.strptime(matches[i]['utcDate'], "%Y-%m-%dT%H:%M:%SZ").strftime('%d-%m-%Y %H:%M'),
                       matches[i]['oddsHomeWin'], matches[i]['oddsDraw'], matches[i]['oddsAwayWin']])
                       #home_teams[i], away_teams[i]])
wb = Workbook()
sheet = wb.active
for item in outputData:
    sheet.append(item)
wb.save('test_PariMatch.xlsx')
wb.close()

headers = outputData[0]
cells = []
for i in range(1, len(outputData)):
    cells.append(outputData[i])

cells = numpy.transpose(cells)
fig = plotly.graph_objs.Figure(data=[plotly.graph_objs.Table(header=dict(values=headers),
                                                             cells=dict(values=cells))])
plotly.offline.plot(fig, filename=dateFrom + '.html')

print('Done')

table = """<table border="0" align="center" width="100%">
    <thead><tr align="center"><th align="center" width="15%">""" + outputData[0][0] \
        + """</th><th align="center" width="15%">""" + outputData[0][1] \
        + """</th><th align="center" width="25%">""" + outputData[0][2] \
        + """</th><th align="center" width="15%">""" + outputData[0][3] \
        + """</th><th align="center" width="15%">""" + outputData[0][4] \
        + """</th><th align="center" width="15%">""" + outputData[0][5] + """</th></tr></thead>
    <tbody>"""

for i in range(1, len(outputData)):
    table += "<tr>"
    for j in range(len(outputData[0])):
        if outputData[i][j] is not None:
            table += """<td align="center">""" + str(outputData[i][j]) + "</td>"
        else:
            table += """<td align="center">-</td>"""
    table += "</tr>"

table += """"</tbody>
             </table>"""

pdf = HTML2PDF()
pdf.add_page(orientation='H')
pdf.write_html(table)
pdf.output(dateFrom + '.pdf')
