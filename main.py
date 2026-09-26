import openpyxl
from datetime import datetime


# Open Excel file
file_name = "1234567.xlsx"

workbook = openpyxl.load_workbook(file_name, data_only=True)
sheet = workbook["Daily Log"]

print("Excel file successfully opened!")
print("Sheet:", sheet.title)


# Variables
valid_days = 0
invalid_days = 0

total_sleep = 0
total_fitness = 0
total_study = 0
total_coding = 0
total_class = 0
total_other = 0

total_tracked = 0
total_free = 0

total_experience_score = 0


# Score values
feeling_scores = {
    "Excellent": 5,
    "Good": 4,
    "Neutral": 3,
    "Low": 2,
    "Stressed": 1
}

satisfaction_scores = {
    "Very Satisfied": 5,
    "Satisfied": 4,
    "Neutral": 3,
    "Unsatisfied": 2,
    "Very Unsatisfied": 1
}

energy_scores = {
    "High": 3,
    "Medium": 2,
    "Low": 1
}


# Read daily records
for row in range(6, 46):

    date = sheet.cell(row, 1).value

    sleep = sheet.cell(row, 2).value
    fitness = sheet.cell(row, 3).value
    study = sheet.cell(row, 4).value
    coding = sheet.cell(row, 5).value
    class_time = sheet.cell(row, 6).value
    other = sheet.cell(row, 8).value

    feeling = sheet.cell(row, 11).value
    satisfaction = sheet.cell(row, 12).value
    energy = sheet.cell(row, 13).value

    # Skip empty rows
    if date is None:
        continue

    # Check activity values
    numeric_values = [
        sleep,
        fitness,
        study,
        coding,
        class_time,
        other
    ]

    valid_numeric_data = True

    for value in numeric_values:

        if not isinstance(value, (int, float)):
            valid_numeric_data = False

        elif value < 0:
            valid_numeric_data = False

    # Check feeling, satisfaction and energy
    valid_experience_data = (
        feeling in feeling_scores
        and satisfaction in satisfaction_scores
        and energy in energy_scores
    )

    # Skip invalid record
    if not valid_numeric_data or not valid_experience_data:
        invalid_days += 1
        continue

    valid_days += 1

    # Add activity totals
    total_sleep += sleep
    total_fitness += fitness
    total_study += study
    total_coding += coding
    total_class += class_time
    total_other += other

    # Calculate total tracked time
    tracked = (
        fitness
        + study
        + coding
        + class_time
        + other
    )

    total_tracked += tracked

    # Calculate free time
    free_time = 1440 - sleep - tracked

    total_free += free_time

    # Convert scores
    feeling_score = feeling_scores[feeling]
    satisfaction_score = satisfaction_scores[satisfaction]
    energy_score = energy_scores[energy]

    # Daily experience score
    daily_experience = (
        feeling_score
        + satisfaction_score
        + energy_score
    ) / 3

    total_experience_score += daily_experience


# Expected days
start_date = datetime(2026, 8, 13)
end_date = datetime(2026, 9, 21)

expected_days = (end_date - start_date).days + 1


# Missing days
missing_days = expected_days - valid_days - invalid_days


# Daily averages
if valid_days > 0:

    average_sleep = total_sleep / valid_days
    average_fitness = total_fitness / valid_days
    average_study = total_study / valid_days
    average_coding = total_coding / valid_days
    average_class = total_class / valid_days
    average_other = total_other / valid_days
    average_tracked = total_tracked / valid_days
    average_free = total_free / valid_days
    average_experience = total_experience_score / valid_days

else:

    average_sleep = 0
    average_fitness = 0
    average_study = 0
    average_coding = 0
    average_class = 0
    average_other = 0
    average_tracked = 0
    average_free = 0
    average_experience = 0


# Calculate indexes
TPI = average_coding

AAI = average_study + average_class

PhAI = average_fitness

SRI = average_sleep

ABI = average_free

TUI = average_tracked

EI = average_experience


# Data Continuity Index
if expected_days > 0:
    DCI = (valid_days / expected_days) * 100
else:
    DCI = 0


# Personal Activity Index
PAI = (
    0.15 * TPI
    + 0.20 * AAI
    + 0.15 * PhAI
    + 0.20 * SRI
    + 0.15 * TUI
    + 0.10 * EI
    + 0.05 * DCI
)


# Data summary
print("\n")
print("----- DATA SUMMARY -----")

print("Expected Days:", expected_days)
print("Valid Days:", valid_days)
print("Missing Days:", missing_days)
print("Invalid / Excluded Records:", invalid_days)

print("\nTotal Sleep:", total_sleep, "minutes")
print("Total Fitness:", total_fitness, "minutes")
print("Total Study:", total_study, "minutes")
print("Total Coding:", total_coding, "minutes")
print("Total Class:", total_class, "minutes")
print("Total Other Activities:", total_other, "minutes")
print("Total Tracked:", total_tracked, "minutes")
print("Total Free / Unaccounted:", total_free, "minutes")


# Daily averages
print("\n")
print("----- DAILY AVERAGES -----")

print("Average Sleep:", round(average_sleep, 2), "min/day")
print("Average Fitness:", round(average_fitness, 2), "min/day")
print("Average Study:", round(average_study, 2), "min/day")
print("Average Coding:", round(average_coding, 2), "min/day")
print("Average Class:", round(average_class, 2), "min/day")
print("Average Other Activities:", round(average_other, 2), "min/day")
print("Average Total Tracked:", round(average_tracked, 2), "min/day")
print("Average Free / Unaccounted:", round(average_free, 2), "min/day")


# Index results
print("\n")
print("----- INDEX RESULTS -----")

print("TPI:", round(TPI, 2), "min/day")
print("AAI:", round(AAI, 2), "min/day")
print("PhAI:", round(PhAI, 2), "min/day")
print("SRI:", round(SRI, 2), "min/day")
print("ABI:", round(ABI, 2), "min/day")
print("TUI:", round(TUI, 2), "min/day")
print("EI:", round(EI, 2), "/ 5")
print("DCI:", round(DCI, 2), "%")


# Final PAI
print("\n")
print("----- PERSONAL ACTIVITY INDEX -----")
print("PAI:", round(PAI, 2))


# Key findings

# Sleep vs Energy
sleep_energy_data = []

for row in range(6, 46):

    date = sheet.cell(row, 1).value
    sleep = sheet.cell(row, 2).value
    energy = sheet.cell(row, 13).value

    if date is not None and energy in energy_scores:

        energy_score = energy_scores[energy]

        sleep_energy_data.append(
            (sleep, energy_score)
        )


print("\n----- SLEEP vs ENERGY -----")

for energy_level, score in energy_scores.items():

    sleep_values = []

    for sleep, energy_score in sleep_energy_data:

        if energy_score == score:
            sleep_values.append(sleep)

    if len(sleep_values) > 0:

        average = sum(sleep_values) / len(sleep_values)

        print(
            energy_level,
            ":",
            round(average, 2),
            "min/day",
            "| Days:",
            len(sleep_values)
        )


# Study vs Satisfaction
study_satisfaction_data = []

for row in range(6, 46):

    date = sheet.cell(row, 1).value
    study = sheet.cell(row, 4).value
    satisfaction = sheet.cell(row, 12).value

    if date is not None and satisfaction in satisfaction_scores:

        satisfaction_score = satisfaction_scores[satisfaction]

        study_satisfaction_data.append(
            (study, satisfaction_score)
        )


print("\n----- STUDY vs SATISFACTION -----")

for satisfaction_level, score in satisfaction_scores.items():

    study_values = []

    for study, satisfaction_score in study_satisfaction_data:

        if satisfaction_score == score:
            study_values.append(study)

    if len(study_values) > 0:

        average = sum(study_values) / len(study_values)

        print(
            satisfaction_level,
            ":",
            round(average, 2),
            "min/day",
            "| Days:",
            len(study_values)
        )


# Coding vs Energy
coding_energy_data = []

for row in range(6, 46):

    date = sheet.cell(row, 1).value
    coding = sheet.cell(row, 5).value
    energy = sheet.cell(row, 13).value

    if date is not None and energy in energy_scores:

        energy_score = energy_scores[energy]

        coding_energy_data.append(
            (coding, energy_score)
        )


print("\n----- CODING vs ENERGY -----")

for energy_level, score in energy_scores.items():

    coding_values = []

    for coding, energy_score in coding_energy_data:

        if energy_score == score:
            coding_values.append(coding)

    if len(coding_values) > 0:

        average = sum(coding_values) / len(coding_values)

        print(
            energy_level,
            ":",
            round(average, 2),
            "min/day",
            "| Days:",
            len(coding_values)
        )