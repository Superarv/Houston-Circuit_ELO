from openpyxl.workbook import Workbook
from openpyxl import load_workbook

teamCodeDict = {
    "0":"Cypress Woods HH", 
    "1":"Eisenhower Senior LR", 
    "2":"Eisenhower Senior AH", 
    "3":"Friendswood PS", 
    "4":"Houston MacArthur CH", 
    "5":"Woodlands College Park GN", 
    "6":"Woodlands College Park MS", 
    "7":"Woodlands College Park MR", 
    "8":"Woodlands College Park OS",
    "9":"Woodlands College Park LO",
    "10":"Barbers Hill BR",
    "11":"Elkins PS",
    "12":"Heights DW",
    "13":"Heights EL",
    "14":"Heights FG",
    "15":"Heights GP",
    "16":"Katy Taylor WW",
    "17":"Memorial CS",
    "18":"Memorial GH",
    "19":"Stephen F Austin JM",
    "20":"Langham Creek BB",
    "21":"Langham Creek CC",
    "22":"Pasadena DD",
    "23":"Kinkaid AK",
    "24":"Kinkaid AS",
    "25":"Caney Creek CI",
    "26":"Caney Creek SW",
    "27":"Woodlands College Park FO",
    "28":"Tomball CP",
    "29":"Tomball Memorial KM",
    "30":"Dulles MS",
    "31":"Dulles MN",
    "32":"Dulles SY",
    "33":"Elkins JJ",
    "34":"Elkins GK",
    "35":"LV Hightower NV",
    "36":"LV Hightower AP",
    "37":"LV Hightower JP",
    "38":"Stephen F Austin KS",
    "39": "Dulles CI",
    "40":"Glenda Dawson JP",
    "41":"Glenda Dawson LS",
    "42":"Glenda Dawson MY",
    "43":"Houston MacArthur GH",
    "44":"Langham Creek BN",
    "45":"Langham Creek EY",
    "46":"Morton Ranch CR",
    "47":"Tomball Memorial MT",
    "48":"Memorial TT",
    "49":"Morton Ranch MN",
    "50":"Stephen F Austin HM",
    "51":"Westside DD",
    "52":"Westide GV",
    "53":"Langham Creek BH",
    "54":"Langham Creek ON",
    "55":"Langham Creek TR",
    "56":"Cypress Woods LR",
    "57":"Cypress Woods AT",
    "58":"Cypress Woods BC",
    "59":"Cypress Woods GC",
    "60":"Pasadena DY",
    "61":"Langham Creek KW",
    "62":"Heights PR"
    }

def getTeamFullName(team):
    team = str(team)
    return teamCodeDict[team]

def eloSetter(elo_A, elo_B, winner): # takes 2 ints and a bool
    putDown_A = 0.05*(elo_A)
    putDown_B = 0.05*(elo_B)

    diff = abs(0.1*(elo_A-elo_B))

    if elo_A > elo_B:
        if winner == True: # expect scenario
            if diff >= putDown_B: # if the gap in elos is significant
                new_A = elo_A + int(0.1*(putDown_A))
                new_B = elo_B - int(0.1*(putDown_B))

            else:
                new_A = elo_A + (putDown_A - diff)
                new_B = elo_B - (putDown_B - diff) 
        if winner == False: # upset scenario
            new_A = elo_A - putDown_A - diff
            new_B = elo_B + putDown_B + diff
    
    
    if elo_B > elo_A:
        if winner == True: # upset scenario
            new_B = elo_B - putDown_B - diff
            new_A = elo_A + putDown_A + diff
        if winner == False: # expect scenario
            if diff >= putDown_A: # if the gap in elos is significant
                new_B = elo_B + int(0.1*(putDown_B))
                new_A = elo_A - int(0.1*(putDown_A))

            else:
                new_B = elo_B + (putDown_B - diff)
                new_A = elo_A - (putDown_A - diff)
    
    if elo_A == elo_B: # if initial elos are equal
        if winner == True:
            new_A = elo_A + putDown_A
            new_B = elo_B - putDown_B
        if winner == False:
            new_A = elo_A - putDown_A
            new_B = elo_B + putDown_B

    return (new_A, new_B)

def printRoundSummary(list):
    for item in list:
        print(item)

def powerRanking(list):
    topTenElos = []
    dup_list = list[:]
    topTenTeams = []
    indexes = []
    for i in range(len(list)):
        topTenElos.append(max(list))
        list.remove(max(list))

    for elo in topTenElos:
        topTenTeams.append(getTeamFullName(dup_list.index(elo)))

    for team in topTenTeams:
        print(f'{topTenTeams.index(team) + 1}.) {team} ({topTenElos[topTenTeams.index(team)]})')

    

    #for team in topTen:
        #print(f'{topTen.index(team) + 1}.) {getTeamFullName(team)}')


## Code to connect excel sheet ##

wb = Workbook()
wb = load_workbook('All-Rounds_DebateElo.xlsx')
ws = wb['Sheet1']

# Number of rounds in current season #

rounds = 0
for row in ws:
    if not all([cell.value == None for cell in row]):
        rounds += 1
rounds -= 1

elos = [1000 for i in range(63)]
#print(elos)

def printTeamRecord(teamName):
    teamCode = list(teamCodeDict.values()).index(teamName)

    aff_losses, aff_wins, neg_losses, neg_wins, elim = 0,0,0,0,0

    for round in range(1, rounds+1):
        
        if teamCode == int(ws.cell(row=round+1,column=3).value):
            if teamCode == int(ws.cell(row=round+1,column=5).value):
                aff_wins += 1
            else:
                aff_losses += 1
            if int(ws.cell(row=round+1,column=6).value) > 0:
                elim += 1
        if teamCode == int(ws.cell(row=round+1,column=4).value):
            if teamCode == int(ws.cell(row=round+1,column=5).value):
                neg_wins += 1
            else:
                neg_losses += 1
            if int(ws.cell(row=round+1,column=6).value) > 0:
                elim += 1

        aff_rounds = aff_losses + aff_wins
        neg_rounds = neg_losses + neg_wins

    print(f'{teamName}\n\nOverall Record : {aff_wins + neg_wins}-{aff_losses + neg_losses}\nAff Record: {aff_wins}-{aff_losses}\nNeg Record: {neg_wins}-{neg_losses}\n{elim} Elimination Rounds')
## Calculate final elos ##

round_summaries = ["[Tourney] [Round] : [AFF] vs. [NEG] --> [WINNER] [FINAL ELOS]"]

for round in range(1, rounds+1):
    team_aff = int(ws.cell(row=round+1,column=3).value)
    team_neg = int(ws.cell(row=round+1,column=4).value)
    team_winner = int(ws.cell(row=round+1,column=5).value)
    tourney = ws.cell(row=round+1,column=1).value
    roundLabel = ws.cell(row=round+1,column=2).value

    aff_elo = elos[team_aff]
    neg_elo = elos[team_neg]
    
    if team_winner == team_aff:
        winnerPC = True
    if team_winner == team_neg:
        winnerPC = False

    new_aff_elo, new_neg_elo = eloSetter(aff_elo, neg_elo, winnerPC)
    
    elos[team_aff] = new_aff_elo
    elos[team_neg] = new_neg_elo

    aff_FULL = getTeamFullName(team_aff)
    neg_FULL = getTeamFullName(team_neg)
    winner_FULL = getTeamFullName(team_winner)

    round_summaries.append(f'{tourney} {roundLabel} : {aff_FULL} ({aff_elo}) vs. {neg_FULL} ({neg_elo}) --> {winner_FULL} [AFF: ({new_aff_elo}) NEG: ({new_neg_elo})]')

#powerRanking(elos)
#printRoundSummary(round_summaries)
while True:
    user = input("\nranking, team, season-summary: ")

    if user == "ranking":
        powerRanking(elos)
    if user == "team":
        printTeamRecord(input("Enter team code: "))
    if user == "season-summary":
        printRoundSummary(round_summaries)
    if user == "exit":
        exit()
