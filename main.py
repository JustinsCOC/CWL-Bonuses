import openpyxl

# Load the overall sheet containing all attacker names
workbook = openpyxl.load_workbook('C:\\Users\\justi\\Downloads\\psychoblue2.xlsx')
overall_sheet = workbook.worksheets[0]

# Create a dictionary to store calculated data for each attacker
attacker_data = {}

# Create a dictionary to store total offense scores for each attacker
attacker_offense = {}

# Create a dictionary to store net stars for each attacker
attacker_net_stars = {}

# Iterate through each round sheet (2 to 8)
for round_number in range(1, 8):
    round_sheet = workbook.worksheets[round_number]

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
        attacker_offense_score = float(attacker_offense_score) if attacker_offense_score else 1

        # Calculate the difference and update the attacker's data
        diff = (defender_townhall - attacker_townhall) - defender_stars
        if attacker_name in attacker_data:
            attacker_data[attacker_name].append(diff)
            attacker_offense[attacker_name] += attacker_offense_score
        else:
            attacker_data[attacker_name] = [diff]
            attacker_offense[attacker_name] = attacker_offense_score

# Calculate net stars and store them in the attacker_net_stars dictionary
for attacker_name in attacker_data:
    total_value = sum(attacker_data[attacker_name])
    total_offense = attacker_offense[attacker_name]
    net_stars = total_value + total_offense
    attacker_net_stars[attacker_name] = net_stars

# Sort the attacker_net_stars dictionary by values (net stars)
sorted_attackers = sorted(attacker_net_stars.items(), key=lambda x: x[1], reverse=True)

# Iterate through the sorted dictionary to display the net stars for each attacker
for attacker_name, net_stars in sorted_attackers:
    print(f"{attacker_name}: Net Stars: {net_stars}")

# Save the updated workbook if needed
# workbook.save('updated_file.xlsx')
