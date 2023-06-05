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

            #Calculation of SER,SEA and SET
            data['SER'] = -10 * data['log_A']
            data['SEA'] = -10 * data['log_B']

            #Calculation of Total SE ie SET
            data['SET'] = data['SER']+data['SEA']

            #Check is SEA is greater than 10dB or not

            
            # Plot SER versus Frequency
            plt.figure(figsize=(8, 6))
            plt.plot(data['Frequency'], data['SER'])
            plt.xlabel('Frequency')
            plt.ylabel('SER')
            plt.title('Frequency versus SER')
            rotation_image_path = os.path.join(directory, f"{file_name.split('.')[0]}_SER.png")
            plt.savefig(rotation_image_path)  # Save the plot as an image
            plt.show()

            # Plot SEA versus Frequency
            plt.figure(figsize=(8, 6))
            plt.plot(data['Frequency'], data['SEA'])
            plt.xlabel('Frequency')
            plt.ylabel('SEA')
            plt.title('Frequency versus SEA')
            ellipticity_image_path = os.path.join(directory, f"{file_name.split('.')[0]}_SEA.png")
            plt.savefig(ellipticity_image_path)  # Save the plot as an image
            plt.show()

            # Plot SET versus Frequency
            plt.figure(figsize=(8, 6))
            plt.plot(data['Frequency'], data['SET'])
            plt.xlabel('Frequency')
            plt.ylabel('SET')
            plt.title('Frequency versus SET')
            ellipticity_image_path = os.path.join(directory, f"{file_name.split('.')[0]}_SET.png")
            plt.savefig(ellipticity_image_path)  # Save the plot as an image
            plt.show()

            # Save the data as an Excel file
            data[['Frequency', 'SER', 'SEA','SET']].to_excel(output_file_path, index=False)


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