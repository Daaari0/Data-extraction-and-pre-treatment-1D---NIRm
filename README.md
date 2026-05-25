# Data Extractor and Plotter for Data stored as .xls by Oscilloscope Promax OD-606

This script automates the extraction, averaging the signal, and visualization of experimental data stored in `.xls` spreadsheet formats. It computes average signal values and standard deviations, calculates elapsed experimental time, and exports results into structured text reports while generating interactive plots.

---

## Required File Naming Format

To gather groups and variables correctly, all `.xls` data files must follow this specific underscore-separated naming convention:

`[X-Value]_[GroupLabel]_[AnythingElse].xls`

* **X-Value:** The independent variable (e.g., absorbance, time step, or numerical index). Can be alphanumeric, but only numeric values are included in final scatter plots.
* **GroupLabel:** A string label used to categorize and color-code data series on the final plot.

> **Example:** `0.5_Control_Run1.xls`, `1.2_SampleB_Trial2.xls`

---

## Features

* **Legacy Compatibility:** Uses `xlrd` to read data points from column C (index 2) starting at row 11 (index 10) of old Excel sheets.
* **Timestamp Calculation:** Extracts the baseline experiment time from cell B1 of each file and converts subsequent entries into elapsed seconds.
* **Statistical Summary:** Automatically computes the mean and standard deviation for each file's data block.
* **Transformed Axis Options:** Offers interactive terminal prompts to plot raw X-values directly or scale them using the absorption factor formula: **1 - 10^(-x)**
* **Clean Data Export:** Outputs a tab-separated text file containing timestamps, deltas, filenames, averages, and deviations.

---

## Prerequisites

Ensure you have Python installed along with the required dependencies:

```bash
pip install xlrd matplotlib
```

## Usage
Run the script.

A graphical folder dialog will appear. Select the directory containing your source .xls files.

Select the destination directory where the final compiled report ([FolderName]_output.txt) should be saved.

Respond to the command-line prompts to choose your preferred X-axis scaling before the script generates and displays the visual scatter plot.
