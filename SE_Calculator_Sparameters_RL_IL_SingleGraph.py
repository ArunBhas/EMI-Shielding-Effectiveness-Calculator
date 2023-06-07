import math
import os
import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog

def process_text_files(directory):
    files = os.listdir(directory)

    for file_name in files:
        if file_name.endswith(".txt"):
            file_path = os.path.join(directory, file_name)
            output_file_path = os.path.join(directory, f"{file_name.split('.')[0]}.xlsx")

            # Perform calculations directly on the text file without reading the lines
            with open(file_path, "r") as file:
                lines = file.readlines()[6:]  # Skip the first 6 lines

            # Convert the lines to DataFrame and perform calculations
            data = pd.DataFrame([line.strip().split() for line in lines],
                                columns=['Frequency', 's11r', 's11i', 's21r', 's21i', 's12r', 's12i', 's22r', 's22i'])
            data = data.astype(float)

            data['s11_mag'] = data['s11r'] ** 2 + data['s11i'] ** 2
            data['s21_mag'] = data['s21r'] ** 2 + data['s21i'] ** 2
            data['s12_mag'] = data['s12r'] ** 2 + data['s12i'] ** 2
            data['s22_mag'] = data['s22r'] ** 2 + data['s22i'] ** 2
            
            # Reflectance and Transmittance calculations
            data['R'] = data['s11_mag']
            data['T'] = data['s21_mag']

            # (1-R)=A and T/(1-R)=B calculations
            data['A'] = 1-data['R']
            data['B'] = data['T']/data['A']

            #log of A and B
            data['log_A'] = data['A'].apply(lambda x: math.log10(x))
            data['log_B'] = data['B'].apply(lambda x: math.log10(x))

            #Calculation of SER,SEA and SET (calculation in terms of 20 log voltage ratio)
            data['SER'] = -10 * data['log_A']
            data['SEA'] = -10 * data['log_B']

            #Calculation of Total Shielding Effectiveness ie SET
            data['SET'] = data['SER']+data['SEA']

            #Check is SEA is greater than 10dB or not

            # Return Loss and Insertion Loss calculations

            data['Return_Loss'] = -10 * data['R'].apply(lambda x: math.log10(x))
            data['Insertion_Loss'] = -10 * data['T'].apply(lambda x: math.log10(x))
            
            # Create a single figure
            plt.figure(figsize=(8, 6))

            # Plot SER versus Frequency
            plt.plot(data['Frequency'], data['SER'], label='SER')

            # Plot SEA versus Frequency
            plt.plot(data['Frequency'], data['SEA'], label='SEA')

            # Plot SET versus Frequency
            plt.plot(data['Frequency'], data['SET'], label='SET')
            # Plot Return Loss versus Frequency
            plt.plot(data['Frequency'], data['Return_Loss'], label='SET')

            plt.xlabel('Frequency (GHz)')
            plt.ylabel('Shielding Effectiveness (dB)')
            plt.title('Frequency versus SER,SEA and SET')
            plt.legend()

            # Save the plot as an image
            image_path = os.path.join(directory, f"{file_name.split('.')[0]}_Graph.png")
            plt.savefig(image_path)

            # Show the plot
            plt.show()

            # Save the data as an Excel file
            data[['Frequency','s11r','s11i','s21r','s21i','s12r','s12i','s22r','s22i','Return_Loss','Insertion_Loss','SER', 'SEA','SET']].to_excel(output_file_path, index=False)


# Create a Tkinter root window
root = tk.Tk()
root.withdraw()

# Ask the user to select a folder
folder_path = filedialog.askdirectory(title="Select Folder")

# Check if a folder was selected
if folder_path:
    # Process text files in the folder
    process_text_files(folder_path)
else:
    print("No folder selected.")