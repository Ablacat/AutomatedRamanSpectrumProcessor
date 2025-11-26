# Automated Raman Spectrum Background Subtraction Processor
# Written by Gabe Kantor 11/26/2025
# Adapted from Vodinh et al. (2006) lowest-point background subtraction method

import numpy as np
from scipy.signal import savgol_filter
import tkinter as tk
from tkinter import filedialog
import os
import pyperclip


def remove_fluo_spectra_lowest_point(input_spectrum, int_width, need_to_be_smoothed=True):
    input_spectrum = np.asarray(input_spectrum)
    if input_spectrum.ndim != 2 or input_spectrum.shape[1] != 2:
        raise ValueError("input_spectrum must be an (N,2) array: [wavenumber, intensity].")

    int_width = int(int_width)
    if int_width < 2:
        raise ValueError("int_width must be >= 2.")
    if int_width % 2 != 0:
        raise ValueError("int_width must be an even integer.")

    x = input_spectrum[:, 0]
    y = input_spectrum[:, 1].copy()

    # Optional smoothing
    preprocessed_y = y.copy()
    if need_to_be_smoothed:
        span = 5
        poly_degree = 1
        preprocessed_y = savgol_filter(y, window_length=span, polyorder=poly_degree)

    n = len(preprocessed_y)
    half_width = int_width // 2
    bg = preprocessed_y.copy()

    # Compute lowest-point background
    for i in range(half_width, n - half_width):
        lowest_intensity = np.inf
        for j in range(1, half_width + 1):
            for k in range(1, half_width + 1):
                temp_intensity = (
                    (preprocessed_y[i + k] - preprocessed_y[i - j])
                    * (j / (k + j))
                    + preprocessed_y[i - j]
                )
                if temp_intensity < lowest_intensity:
                    lowest_intensity = temp_intensity
        bg[i] = lowest_intensity

    clean_y = preprocessed_y - bg
    clean_spectrum = np.column_stack((x, clean_y))
    return clean_spectrum


def process_one_file():
    # Set up Tkinter dialogs
    root = tk.Tk()
    root.withdraw()

    # Select CSV input file
    filename = filedialog.askopenfilename(
        title="Select spectrum CSV file",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not filename:
        print("No file selected. Returning to menu.\n")
        root.destroy()
        return

    # Load CSV
    input_spectrum = np.loadtxt(filename, delimiter=",")

    # Parameters
    b_need_to_be_smoothed = True
    int_width = 30

    clean_spectrum = remove_fluo_spectra_lowest_point(
        input_spectrum,
        int_width=int_width,
        need_to_be_smoothed=b_need_to_be_smoothed
    )

    # Output CSV dialog
    default_name = os.path.splitext(os.path.basename(filename))[0] + "_bkgrndsub.csv"
    save_path = filedialog.asksaveasfilename(
        title="Save background-subtracted CSV",
        defaultextension=".csv",
        initialfile=default_name,
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )

    # Save as TAB-delimited
    if save_path:
        np.savetxt(save_path, clean_spectrum, delimiter="\t", fmt="%.8g")
        print(f"Saved processed spectrum to: {save_path}")
    else:
        print("No output file selected (skipping save).")

    # Copy to clipboard
    lines = ["\t".join(f"{v:.8g}" for v in row) for row in clean_spectrum]
    clip_text = "\n".join(lines)

    try:
        pyperclip.copy(clip_text)
        print("Clean spectrum copied to clipboard (tab-delimited for Excel).\n")
    except pyperclip.PyperclipException as e:
        print("Could not copy to clipboard:", e, "\n")

    root.destroy()


if __name__ == "__main__":
    while True:
        process_one_file()
        ans = input("Process another file? (y/n): ").strip().lower()
        if ans != "y":
            print("Session ended.")
            break