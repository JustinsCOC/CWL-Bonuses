import openpyxl
from colorama import Fore, Style

# Load the overall sheet containing all attacker names
workbook = openpyxl.load_workbook('C:\\Users\\justi\\Downloads\\pblue test.xlsx')
overall_sheet = workbook.worksheets[0]

# Create a dictionary to store calculated data for each attacker
attacker_data = {}

# Iterate through each round sheet (3 to 9)
for round_number in range(3, 10):
    round_sheet = workbook.worksheets[round_number]

    # Create a dictionary to store total offense scores for each attacker in this round
    attacker_offense_round = {}

    # Create a dictionary to store scores for each attacker in this round
    attacker_score_round = {}

    # Create a dictionary to store gross stars for each attacker in this round
    attacker_gross_stars = {}

    # Iterate through each row in the round sheet
    for row in round_sheet.iter_rows(min_row=2, values_only=True):
        attacker_name = row[2]  # Column C (attacker's name)
        attacker_townhall = row[11]  # Column L (attacker's town hall)
        defender_townhall = row[13]  # Column N (defender's town hall)
        defender_stars = row[14]  # Column O (attacker's defense stars)
        attacker_offense_score = row[4]  # Column E (attacker's offense score)

        # Handle data type conversion and None values
        if attacker_name is None:
            continue
        attacker_townhall = float(attacker_townhall) if attacker_townhall else 0
        defender_townhall = float(defender_townhall) if defender_townhall else 0
        defender_stars = float(defender_stars) if defender_stars else 2
        attacker_offense_score = float(attacker_offense_score) if attacker_offense_score else 0

        # Calculate total offense and update the attacker's data for this round
        if attacker_name in attacker_offense_round:
            attacker_offense_round[attacker_name] += attacker_offense_score
        else:
            attacker_offense_round[attacker_name] = attacker_offense_score

        # Calculate score based on new criteria
        if defender_townhall < attacker_townhall:
            townhall_difference = -1
        elif defender_townhall > attacker_townhall:
            townhall_difference = 1
        else:
            townhall_difference = 0  # No change in score if townhalls are equal

        # Calculate net score including defensive stars
        score = townhall_difference - defender_stars

        # Update the attacker's data for this round
        if attacker_name in attacker_score_round:
            attacker_score_round[attacker_name].append(score)
        else:
            attacker_score_round[attacker_name] = [score]

        # Calculate gross stars (total offensive stars)
        if attacker_name in attacker_gross_stars:
            attacker_gross_stars[attacker_name] += attacker_offense_score
        else:
            attacker_gross_stars[attacker_name] = attacker_offense_score

    # Update the attacker_data dictionary with total offense, scores, and gross stars for this round
    for attacker_name, offense_score in attacker_offense_round.items():
        if attacker_name in attacker_data:
            attacker_data[attacker_name]['offense'].append(offense_score)
            attacker_data[attacker_name]['gross_stars'].append(attacker_gross_stars[attacker_name])
            if attacker_name in attacker_score_round:
                attacker_data[attacker_name]['scores'].append(attacker_score_round[attacker_name])
        else:
            attacker_data[attacker_name] = {'offense': [offense_score],
                                            'scores': [attacker_score_round.get(attacker_name, [])],
                                            'gross_stars': [attacker_gross_stars[attacker_name]]}

# Calculate the net stars and gross stars for each player and store them in dictionaries
attacker_net_stars = {}
attacker_gross_stars = {}
for attacker_name, data in attacker_data.items():
    total_scores = sum(sum(round_scores) for round_scores in data['scores'])
    total_offense = sum(data['offense'])
    net_stars = total_scores + total_offense
    attacker_net_stars[attacker_name] = net_stars
    attacker_gross_stars[attacker_name] = sum(data['gross_stars'])

# Sort the attacker_net_stars dictionary by values (net stars)
sorted_attackers = sorted(attacker_net_stars.items(), key=lambda x: x[1], reverse=True)

# Print formatted output with colors and style
print("\nNet Stars Leaderboard:\n")
print("{:<8} {:<30} {} {}".format("Rank", "Attacker Name", "Net Stars", "Gross Stars"))
print("=" * 65)
for rank, (attacker_name, net_stars) in enumerate(sorted_attackers, start=1):
    if net_stars >= 0:
        color = Fore.GREEN
    else:
        color = Fore.RED
    gross_stars = attacker_gross_stars.get(attacker_name, 0)
    formatted_output = "{:<8} {:<30} {}{:>10}{}{:>12}".format(rank, attacker_name, color, net_stars, Style.RESET_ALL, gross_stars)
    print(formatted_output)
