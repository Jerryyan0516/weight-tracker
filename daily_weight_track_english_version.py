import pandas as pd
from datetime import datetime
import os
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.dates as mdates
import tkinter as tk
from tkinter import simpledialog, messagebox

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
plt.rcParams['axes.unicode_minus'] = False

file_path = r'C:\Users\65179\Desktop\每日体重记录\date_weight.xlsx'

def record_weight(weight, file_path):
    """
    Record weight data and write it to an Excel file, while comparing with the last weight and providing feedback.
    """
    # Get the current date and time
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Create new data containing the current time and weight
    new_data = pd.DataFrame({'Time': [current_time], 'Weight (kg)': [weight]})

    # Check if the file exists
    if os.path.exists(file_path):
        # If the file exists, read the existing data
        df = pd.read_excel(file_path)

        # Convert the time column to datetime format to avoid data type confusion
        df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
        new_data['Time'] = pd.to_datetime(new_data['Time'], errors='coerce')

        # Get the weight of the last record
        last_weight = df.iloc[-1]['Weight (kg)']

        # Compare weight changes
        weight_diff = weight - last_weight
        if weight_diff > 0:
            message = f"Increased by {weight_diff:.2f} kg since last time, consider increasing your exercise!"
        elif weight_diff < 0:
            message = f"Decreased by {abs(weight_diff):.2f} kg since last time, keep up the good work!"
        else:
            message = "No change in weight, keep maintaining your habits!"

        # Show a popup with the result
        messagebox.showinfo("Weight Change Result", message)

        # Append new data to the existing data
        df = pd.concat([df, new_data], ignore_index=True)
    else:
        # If the file does not exist, create a new DataFrame
        df = new_data

    # Sort data by the time column
    df = df.sort_values(by='Time')

    # Write the sorted data back to the Excel file
    df.to_excel(file_path, index=False)


def plot_weight_trend(file_path):
    """
    Read the Excel file and generate a line chart of weight changes.
    """
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Convert the time column to datetime format
    df['Time'] = pd.to_datetime(df['Time'])

    # Plot the line chart
    plt.figure(figsize=(20, 12))
    plt.plot(df['Time'], df['Weight (kg)'], marker='o', linestyle='-', color='b')

    # Set chart title and axis labels
    plt.title("YY's Weight Change Line Chart", fontsize=18)
    plt.xlabel('Date', fontsize=22)
    plt.ylabel('Weight (kg)', fontsize=22)

    # Set date format
    date_format = mdates.DateFormatter('%Y-%m-%d %H:%M:%S')
    plt.gca().xaxis.set_major_formatter(date_format)

    # Adjust the display of date labels
    plt.gcf().autofmt_xdate()

    # Rotate the x-axis labels for better date display
    plt.xticks(rotation=45)

    # Add grid
    plt.grid(True)

    # Display the chart
    plt.tight_layout()  # Automatically adjust subplot parameters to fill the figure area
    plt.show()


def main():
    """
    Main function to execute weight recording and line chart plotting.
    """
    # Create a Tkinter window
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Pop up an input box for the user to enter weight
    weight_input = simpledialog.askstring("Weight Record", "Please enter today's weight (kg) (or press Enter to view the chart):")

    # If the input is an empty string, display the chart directly
    if weight_input is None or weight_input.strip() == '':
        plot_weight_trend(file_path)
    else:
        weight = float(weight_input)
        record_weight(weight, file_path)
        plot_weight_trend(file_path)


# Call the main function
if __name__ == "__main__":
    main()
