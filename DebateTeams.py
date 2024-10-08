import pandas as pd
import math

file_path = r'C:\Users\deepe\Debate\random_applications.xlsx'
sheet_name = 'Sheet1'

df = pd.read_excel(file_path, sheet_name=sheet_name)

debate_participants = df[df['Role'] == 'Debate'].copy()
judge_participants = df[df['Role'] == 'Judge'].copy()
spectator_participants = df[df['Role'] == 'Spectate (maybe help judging?)'].copy()

def normalize_name(name):
    return ''.join(name.split()).lower()

def form_teams(participants, role):
    teams = []
    paired_names = set()

    for idx, row in participants.iterrows():
        if pd.notna(row['Group With']):
            normalized_group_with = normalize_name(row['Group With'])
            normalized_names = participants['First + Last Name'].apply(normalize_name)

            if normalized_group_with in normalized_names.values:
                partner_idx = participants[normalized_names == normalized_group_with].index[0]
                if partner_idx not in paired_names:
                    teams.append((
                        row['First + Last Name'], row['Skill Level'], 
                        participants.at[partner_idx, 'First + Last Name'], 
                        participants.at[partner_idx, 'Skill Level'],
                        role  
                    ))
                    paired_names.add(idx)
                    paired_names.add(partner_idx)

    unpaired = participants[~participants.index.isin(paired_names)].copy()

    beginners = unpaired[unpaired['Skill Level'] == 'Beginner']
    intermediates = unpaired[unpaired['Skill Level'] == 'Intermediate']
    advanced = unpaired[unpaired['Skill Level'] == 'Advanced']

    combined_list = pd.concat([beginners, intermediates, advanced])

    while len(combined_list) > 1:
        front = combined_list.iloc[0]
        back = combined_list.iloc[-1]

        teams.append((
            front['First + Last Name'], front['Skill Level'], 
            back['First + Last Name'], back['Skill Level'],
            role  
        ))
        
        combined_list = combined_list.iloc[1:-1]

    if len(combined_list) == 1:
        remaining = combined_list.iloc[0]
        teams.append((remaining['First + Last Name'], remaining['Skill Level'], None, None, role))

    return teams

debate_teams = form_teams(debate_participants, 'Debate')
judge_teams = form_teams(judge_participants, 'Judge')

group_count = math.ceil(len(debate_teams) / 4)

spectator_count = len(spectator_participants)
spectators_per_group = spectator_count // group_count
extra_spectators = spectator_count % group_count

team_data = []
spectator_index = 0
judge_index = 0

for group_num in range(group_count):
    current_group = []

    if judge_index < len(judge_teams):
        judge_team = judge_teams[judge_index]
        current_group.append([judge_team[0], judge_team[2], [judge_team[1], judge_team[3]], judge_team[4]])
        judge_index += 1
    else:
        current_group.append([None, None, [None, None], 'Judge'])

    for i in range(4):
        debate_idx = group_num * 4 + i
        if debate_idx < len(debate_teams):
            debate_team = debate_teams[debate_idx]
            current_group.append([debate_team[0], debate_team[2], [debate_team[1], debate_team[3]], debate_team[4]])
        else:
            current_group.append([None, None, [None, None], 'Debate'])

    num_spectators_to_add = spectators_per_group + (1 if group_num < extra_spectators else 0)
    for _ in range(num_spectators_to_add):
        if spectator_index < spectator_count:
            spectator = spectator_participants.iloc[spectator_index]
            current_group.append([spectator['First + Last Name'], None, [spectator['Skill Level']], 'Spectator'])
            spectator_index += 1

    team_data.extend(current_group)
    team_data.append([None, None, None, None]) 

while judge_index < len(judge_teams):
    current_group = []
    judge_team = judge_teams[judge_index]
    current_group.append([judge_team[0], judge_team[2], [judge_team[1], judge_team[3]], judge_team[4]])
    team_data.extend(current_group)
    team_data.append([None, None, None, None])  
    judge_index += 1

team_df = pd.DataFrame(team_data, columns=['Participant 1', 'Participant 2', 'Skill Levels', 'Role'])

with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    team_df.to_excel(writer, sheet_name='Sheet2', index=False)

print(f"Debate teams, judges, and spectators have been saved to 'Sheet2' in {file_path}")
