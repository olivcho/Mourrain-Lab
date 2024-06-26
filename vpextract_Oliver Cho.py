'''
Mourrain Lab - Department of Psychiatry & Behavioral Sciences, Stanford University

Author: Oliver Cho (olivercho007@gmail.com)
Affiliation: The Nueva School, Class of 2025
Date: June 2024

This script automates data analysis for ViewPoint .xlsx files.
The original code for this script was written by Louis Leung, Ph.D. in MATLAB.
Code was rewritten in Python for speed and readability + added new features to analyze sleep bouts.

--------------------------------------------

Requirements: Pandas, NumPy, Matplotlib, Datetime, tkinter, OS.

If you do not have the requirements installed, open your command prompt/terminal and type in "pip install [package]"
'''

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from tkinter import filedialog, Tk, Label, Entry, Button, IntVar, StringVar, Frame, messagebox
import os

file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
df = pd.read_excel(file_path)
df.head()

output_dir = os.path.join(os.path.dirname(file_path), 'Summary')
os.makedirs(output_dir, exist_ok=True)

y_figsize = 15
x_figsize = 12.5

vprawData = pd.DataFrame(np.zeros((len(df), 15)))

startDate = df.iloc[:, 27]
vprawData[0] = startDate

startTime = df.iloc[:, 28]

startTime = pd.to_datetime(startTime, format='%H:%M:%S').dt.time
def total_minutes_elapsed(time):
    base_time = datetime.strptime('18:00:00', '%H:%M:%S').time()
    delta = timedelta(hours=time.hour, minutes=time.minute, seconds=time.second) - timedelta(hours=base_time.hour, minutes=base_time.minute, seconds=base_time.second)
    return delta.seconds // 60
vprawData[1] = startTime.apply(total_minutes_elapsed)

daynightIdx = (vprawData.iloc[:,1] > 300) & (vprawData.iloc[:,1] < 900) # midnight labeled with True
vprawData[2] = daynightIdx
    
vprawData.iloc[:, 3:5] = df.iloc[:, 14:16].to_numpy()
    
animalNum = df.iloc[:, 1].str.extract(r'z(\d+)').astype(float).to_numpy().flatten()
vprawData.iloc[:, 5] = animalNum
    
vprawData.iloc[:, 6:15] = df.iloc[:, 18:27].to_numpy()

vprawData = vprawData.rename(columns={0: 'startDate',
                                      1: 'startTime',
                                      2: 'night',
                                      3: 'start',
                                      4: 'end',
                                      5: 'animalNum',
                                      6: 'freezeCount',
                                      7: 'freezeDuration',
                                      8: 'midCount',
                                      9: 'midDuration',
                                      10: 'burstCount',
                                      11: 'burstDuration',
                                      12: 'zeroCount',
                                      13: 'zeroDuration',
                                      14: 'activityIntegral'})

before_filter = len(vprawData)
vprawData = vprawData[(vprawData['start'] == vprawData['start'].round()) &
                      (vprawData['end'] == vprawData['end'].round()) &
                      (vprawData['start'] != vprawData['end'])]
after_filter = len(vprawData)

def collect_inputs():
    global needTotalBoxBool, removeTransBool, binWidth, sleepwakeMatrixBool, sleepBoutBool, removeBool, startTreat, endTreat

    try:
        needTotalBoxBool = needTotalBoxBool_var.get()
        removeTransBool = removeTransBool_var.get()
        binWidth = int(binWidth_var.get())
        sleepwakeMatrixBool = sleepwakeMatrix_var.get()
        sleepBoutBool = sleepBout_var.get()
        removeBool = removeBool_var.get()
        startTreat = int(startTreat_var.get())
        endTreat = int(endTreat_var.get())

        root.destroy()
    except ValueError:
        messagebox.showerror("Invalid Input", "Please enter valid inputs.")
        return

root = Tk()
root.title("Data Analysis Input")

needTotalBoxBool_var = IntVar(value=1)
removeTransBool_var = IntVar(value=1)
binWidth_var = StringVar(value=10)
sleepwakeMatrix_var = IntVar(value=1)
sleepBout_var = IntVar(value=1)
removeBool_var = StringVar(value=0)
startTreat_var = StringVar(value=0)
endTreat_var = IntVar(value=0)

frame = Frame(root)
frame.pack(pady=20)

Label(frame, text="Do you want TotalBoxPlots (1 is yes): ").grid(row=0, column=0, sticky='w')
Entry(frame, textvariable=needTotalBoxBool_var).grid(row=0, column=1)

Label(frame, text="Do you want a sleep/wake matrix? (1 is yes): ").grid(row=1, column=0, sticky='w')
Entry(frame, textvariable=sleepwakeMatrix_var).grid(row=1, column=1)

Label(frame, text="Do you want a sleep bout duration matrix? (1 is yes): ").grid(row=2, column=0, sticky='w')
Entry(frame, textvariable=sleepBout_var).grid(row=2, column=1)

Label(frame, text="Set time window Bin width (min): ").grid(row=3, column=0, sticky='w')
Entry(frame, textvariable=binWidth_var).grid(row=3, column=1)

Label(frame, text="Do you want to remove the Day/Night transition minute? (1 is yes): ").grid(row=4, column=0, sticky='w')
Entry(frame, textvariable=removeTransBool_var).grid(row=4, column=1)

Label(frame, text="Excluded wells (1-24 // enter space separated): ").grid(row=5, column=0, sticky='w')
Entry(frame, textvariable=removeBool_var).grid(row=5, column=1)

Label(frame, text="Treatment Started: ").grid(row=6, column=0, sticky='w')
Entry(frame, textvariable=startTreat_var).grid(row=6, column=1)

Label(frame, text="Treatment Ended: ").grid(row=7, column=0, sticky='w')
Entry(frame, textvariable=endTreat_var).grid(row=7, column=1)

Button(frame, text="Submit", command=collect_inputs).grid(row=8, columnspan=2, pady=10)

root.mainloop()

num_animals = int(vprawData['animalNum'].max())
noofBins = len(vprawData[vprawData['animalNum'] == 1]) // binWidth

# Remove measurements taken during heatshock or adding drugs
if startTreat or endTreat != 0:
    vprawData.loc[startTreat:endTreat*num_animals, 'freezeCount':'activityIntegral'] = np.nan

# Exclude the minute bin when the light turns on and off
if removeTransBool == 1:
    vprawData.loc[vprawData['startTime'] == 300, 'freezeCount':'activityIntegral'] = np.nan
    vprawData.loc[vprawData['startTime'] == 900, 'freezeCount':'activityIntegral'] = np.nan
    print("Day and Night Removed!")

# Remove any fish that didn't survive the experiment
deleteWells = list(map(int, removeBool.split()))
for well in deleteWells:
    vprawData.loc[vprawData['animalNum'] == well, 'freezeCount':'activityIntegral'] = np.nan

vprawData.sort_values(['animalNum', 'startTime'], ignore_index=True, inplace=True)
vprawData.insert(0, 'animalNum', vprawData.pop('animalNum'))

vprawData.to_excel(os.path.join(output_dir, 'vprawData.xlsx'), index=False)
print(f"vprawData saved to {os.path.join(output_dir)}")



vpMeasure = np.zeros((noofBins, 3, num_animals))

for animalIdx in range(num_animals):

    animal_data = vprawData[vprawData['animalNum'] == animalIdx + 1]

    noofBins = len(animal_data) // binWidth

    for binIdx in range(noofBins):
        binRange = slice(binIdx * binWidth, (binIdx + 1) * binWidth)
        vpMeasure[binIdx, 0, animalIdx] = animal_data.iloc[binRange]['startTime'].iloc[0]
        vpMeasure[binIdx, 1, animalIdx] = animal_data.iloc[binRange]['midDuration'].sum()
        vpMeasure[binIdx, 2, animalIdx] = (animal_data.iloc[binRange]['midDuration'] == 0).sum()



if needTotalBoxBool:
    fig, axs = plt.subplots(4, 6, figsize=(y_figsize, x_figsize))
    fig.suptitle('Activity Plot', fontsize=16)

    axs = axs.flatten()

    for i in range(num_animals):
        axs[i].plot(vpMeasure[:, 0, i], vpMeasure[:, 1, i], color='darkblue')
        axs[i].grid(False)
        
        axs[i].set_ylabel('')
        if i % 6 == 0:
            axs[i].set_ylabel('Abs Activity (sec/10min)', fontsize=14)

        axs[i].set_ylim(0, 150)
        axs[i].set_yticks([0, 50, 100, 150])
        axs[i].set_xticks([])
        axs[i].set_xlabel(f'Well {i + 1}', fontsize=14)

        axs[i].annotate('', xy=(300, 145), xytext=(300, 150),
                        arrowprops=dict(facecolor='black', shrink=0, width=3, headwidth=12))
        
        axs[i].annotate('', xy=(900, 145), xytext=(900, 150),
                        arrowprops=dict(facecolor='white', edgecolor='black', shrink=0, width=3, headwidth=12))

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.savefig(os.path.join(output_dir, 'activity_plot.png'))

if sleepwakeMatrixBool:
   
   fig, axs = plt.subplots(4, 6, figsize=(y_figsize, x_figsize))
   fig.suptitle('Sleep Plot', fontsize=16)

   axs = axs.flatten()

   for i in range(num_animals):
      axs[i].plot(vpMeasure[:, 0, i], vpMeasure[:, 2, i], color='darkred')
      axs[i].grid(False)
        
      axs[i].set_ylabel('')
      if i % 6 == 0:
         axs[i].set_ylabel('Abs Sleep (min/10min)', fontsize=14)

      axs[i].set_ylim(0, 10)
      axs[i].set_yticks([0, 2, 4, 6, 8, 10])
      axs[i].set_xticks([])
      axs[i].set_xlabel(f'Well {i + 1}', fontsize=14)

      axs[i].annotate('', xy=(300, 9.5), xytext=(300, 10),
                        arrowprops=dict(facecolor='black', shrink=0, width=3, headwidth=12))
        
      axs[i].annotate('', xy=(900, 9.5), xytext=(900, 10),
                        arrowprops=dict(facecolor='white', edgecolor='black', shrink=0, width=3, headwidth=12))

   plt.tight_layout(rect=[0, 0.03, 1, 0.95])
   plt.savefig(os.path.join(output_dir, 'sleep_plot.png'))



noofMin = len(vprawData[vprawData['animalNum'] == 1])
vpSleepBout = np.zeros((noofMin, 4, num_animals))

for animalIdx in range(num_animals):

    animal_data = vprawData[vprawData['animalNum'] == animalIdx + 1]

    for binIdx in range(noofMin):
        vpSleepBout[binIdx, 0, animalIdx] = animal_data.iloc[binIdx]['startTime']
        vpSleepBout[binIdx, 1, animalIdx] = animal_data.iloc[binIdx]['freezeCount']
        vpSleepBout[binIdx, 2, animalIdx] = animal_data.iloc[binIdx]['freezeDuration']

        if animal_data.iloc[binIdx]['freezeDuration'].round() == 60 and animal_data.iloc[binIdx]['freezeCount'] == 0:
            vpSleepBout[binIdx, 3, animalIdx] = 1
        else:
            vpSleepBout[binIdx, 3, animalIdx] = 0

if sleepBoutBool:

   fig, axs = plt.subplots(8, 3, figsize=(y_figsize, x_figsize))
   fig.suptitle('Sleep Bout Plot', fontsize=16)

   axs = axs.flatten()

   for i in range(num_animals):
      axs[i].bar(np.arange(len(vpSleepBout[:, 0, i])), 
            vpSleepBout[:, 3, i], 
            color='darkblue', width=0.2, edgecolor='darkblue')
      
      axs[i].grid(False)
        
      axs[i].set_ylabel('')
      if i % 6 == 0:
         axs[i].set_ylabel('Sleep Bouts')

      axs[i].set_yticks([])
      axs[i].set_xticks([])
      axs[i].set_xlabel(f'Well {i + 1}', fontsize=16)

      axs[i].annotate('', xy=(300, 1), xytext=(300, 1.025),
                        arrowprops=dict(facecolor='black', shrink=0, width=3, headwidth=12))
        
      axs[i].annotate('', xy=(900, 1), xytext=(900, 1.025),
                        arrowprops=dict(facecolor='white', edgecolor='black', shrink=0, width=3, headwidth=12))
      
   plt.tight_layout(rect=[0, 0.03, 1, 0.95])
   plt.savefig(os.path.join(output_dir, 'sleep_bout_plot.png'))

# Downloading data
data = []
for animalIdx in range(num_animals):
    for binIdx in range(noofMin):
        row = [
            animalIdx + 1,
            vpSleepBout[binIdx, 0, animalIdx], 
            vpSleepBout[binIdx, 1, animalIdx],
            vpSleepBout[binIdx, 2, animalIdx],
            vpSleepBout[binIdx, 3, animalIdx]
        ]
        data.append(row)

columns = ['animalIdx', 'startTime', 'freezeCount', 'freezeDuration', 'isSleepBout']
vpSleepBout_df = pd.DataFrame(data, columns=columns)
vpSleepBout_df.to_excel(os.path.join(output_dir, 'vpSleepBoutData.xlsx'), index=False)
print(f"vpSleepBoutData saved to {os.path.join(output_dir)}")

if sleepBoutBool:    
    sleep_bout_counts = []

    for animalIdx in range(num_animals):
        total_sleep_bout_count = 0
        total_sleep_time = 0
        day_sleep_bout_count = 0
        day_sleep_time = 0
        night_sleep_bout_count = 0
        night_sleep_time = 0
        is_sleep_bout = False

        for binIdx in range(len(vpSleepBout)):
            is_day = binIdx < 300 or binIdx >= 900
            is_night = 300 <= binIdx < 900

            if vpSleepBout[binIdx, 3, animalIdx] == 1:
                if not is_sleep_bout:
                    total_sleep_bout_count += 1
                    is_sleep_bout = True
                    if is_day:
                        day_sleep_bout_count += 1
                    elif is_night:
                        night_sleep_bout_count += 1
                total_sleep_time += 1
                if is_day:
                    day_sleep_time += 1
                elif is_night:
                    night_sleep_time += 1
            else:
                if is_sleep_bout:
                    is_sleep_bout = False
        
        sleep_bout_counts.append({
            'Fish ID': animalIdx + 1,
            'Day Sleep Bouts': day_sleep_bout_count, 
            'Day Sleep Time (min)': day_sleep_time,
            'Night Sleep Bouts': night_sleep_bout_count,
            'Night Sleep Time (min)': night_sleep_time,
            'Total Sleep Bouts': total_sleep_bout_count,
            'Total Sleep Time (min)': total_sleep_time,
            'Day Average (min)': day_sleep_time / day_sleep_bout_count if day_sleep_bout_count else 0, 
            'Night Average (min)': night_sleep_time / night_sleep_bout_count if night_sleep_bout_count else 0,
            'Total Average (min)': total_sleep_time / total_sleep_bout_count if total_sleep_bout_count else 0
        })

    sleep_bouts_df = pd.DataFrame(sleep_bout_counts)
    sleep_bouts_df.to_excel(os.path.join(output_dir, 'vpSleepBoutTotalData.xlsx'), index=False)
    print(f"vpSleepBoutTotalData saved to {os.path.join(output_dir)}")

    y_min = 0
    y_max = np.max(sleep_bouts_df['Total Sleep Time (min)'])

    plt.figure(figsize=(6, 4))
    plt.suptitle('Sleep Bout Bar Chart', fontsize=16)
    plt.bar(sleep_bouts_df['Fish ID'], sleep_bouts_df['Total Sleep Time (min)'], color='darkblue')
    plt.bar(sleep_bouts_df['Fish ID'], sleep_bouts_df['Total Sleep Bouts'], color='red')
    plt.ylim(y_min, y_max + 300)

    plt.xlabel('Fish ID')
    plt.ylabel('Total Time (minutes)/Frequency', fontsize=7)
    plt.title('Total length of Sleep Bouts + # of Sleep Bouts by Fish/Well', fontsize=10)
    plt.legend(['Total Sleep Time (minutes)', 'Number of Sleep Bouts'])

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.savefig(os.path.join(output_dir, 'total_sleep_bout_statistics.png'))

    y_min = 0
    y_max = np.max(sleep_bouts_df['Total Average (min)'])

    plt.figure(figsize=(6, 4))
    plt.suptitle('Average Sleep Bout Chart', fontsize=16)
    plt.stem(sleep_bouts_df['Fish ID'], sleep_bouts_df['Total Average (min)'], linefmt='darkblue', markerfmt='o', basefmt=' ')
    plt.stem(sleep_bouts_df['Fish ID'], sleep_bouts_df['Night Average (min)'], linefmt='darkgreen', markerfmt='o', basefmt=' ')
    plt.stem(sleep_bouts_df['Fish ID'], sleep_bouts_df['Day Average (min)'], linefmt='darkred', markerfmt='o', basefmt=' ')
    plt.ylim(y_min, y_max + 2)
    plt.xlabel('Fish ID')
    plt.ylabel('Average Sleep Bout Length (min)', fontsize=7)
    plt.legend(['Total Average', 'Night Average', 'Day Average'])
    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.savefig(os.path.join(output_dir, 'average_sleep_bout_statistics.png'))

'''    y_min = 0
    y_max = np.max(sleep_bouts_df['Total Average (min)'])

    plt.figure(figsize=(6, 4))
    plt.suptitle('Average Sleep Bout Bar Chart', fontsize=16)
    plt.bar(sleep_bouts_df['Fish ID'], sleep_bouts_df['Total Average (min)'], color='darkblue')
    plt.bar(sleep_bouts_df['Fish ID'], sleep_bouts_df['Night Average (min)'], color='darkgreen')
    plt.bar(sleep_bouts_df['Fish ID'], sleep_bouts_df['Day Average (min)'], color='darkred')
    plt.ylim(y_min, y_max + 2)

    plt.xlabel('Fish ID')
    plt.ylabel('Average Sleep Bout Length (min)', fontsize=7)
    plt.legend(['Total Average', 'Night Average', 'Day Average'])

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.savefig(os.path.join(output_dir, 'average_sleep_bout_statistics.png'))'''

if sleepBoutBool:
    
    fig, axs = plt.subplots(8, 3, figsize=(y_figsize, x_figsize))
    fig.suptitle('Sleep Bouts w/ Sleep Overlay', fontsize=16)
    
    axs = axs.flatten()
    
    for i in range(num_animals):
        axs[i].bar(np.arange(len(vpSleepBout[:, 0, i])), 
            vpSleepBout[:, 3, i], 
            color='darkblue', width=0.2, edgecolor='darkblue')
        
        axs[i].plot(vpMeasure[:, 0, i], vpMeasure[:, 2, i] / 10, color='red', label='Abs Sleep (min/10min)')
        axs[i].grid(False)
        axs[i].set_ylabel('')
        if i % 6 == 0:
            axs[i].set_ylabel('Sleep Bout (sleep bout/min)')
            
        axs[i].set_yticks([])
        axs[i].set_xticks([])
        axs[i].set_xlabel(f'Well {i + 1}', fontsize=16)

        axs[i].annotate('', xy=(300, 1), xytext=(300, 1.025),
                        arrowprops=dict(facecolor='black', shrink=0, width=3, headwidth=12))
        
        axs[i].annotate('', xy=(900, 1), xytext=(900, 1.025),
                        arrowprops=dict(facecolor='white', edgecolor='black', shrink=0, width=3, headwidth=12))

    plt.tight_layout(rect=[0, 0.03, 1, 0.95])
    plt.savefig(os.path.join(output_dir, 'sleep_bouts_sleep_overlay.png'))

if sleepBoutBool:
   
   fig, axs = plt.subplots(8, 3, figsize=(y_figsize, x_figsize))
   fig.suptitle('Sleep Bouts w/ Activity Overlay', fontsize=16)

   axs = axs.flatten()

   for i in range(num_animals):
      axs[i].bar(np.arange(len(vpSleepBout[:, 0, i])), 
            vpSleepBout[:, 3, i], 
            color='darkblue', width=0.2, edgecolor='darkblue')
      
      axs[i].plot(vpMeasure[:, 0, i], vpMeasure[:, 1, i] / np.max(vpMeasure[:, 1, :]), color='red')

      axs[i].grid(False)
        
      axs[i].set_ylabel('')
      if i % 6 == 0:
         axs[i].set_ylabel('Sleep Bout (sleep bout/min)')

      axs[i].set_yticks([])
      axs[i].set_xticks([])
      axs[i].set_xlabel(f'Well {i + 1}', fontsize=16)

      axs[i].annotate('', xy=(300, 1), xytext=(300, 1.025),
                        arrowprops=dict(facecolor='black', shrink=0, width=3, headwidth=12))
        
      axs[i].annotate('', xy=(900, 1), xytext=(900, 1.025),
                        arrowprops=dict(facecolor='white', edgecolor='black', shrink=0, width=3, headwidth=12))

   plt.tight_layout(rect=[0, 0.03, 1, 0.95])
   plt.savefig(os.path.join(output_dir, 'sleep_bouts_activity_overlay.png'))

plt.show()

print("Script Complete!")