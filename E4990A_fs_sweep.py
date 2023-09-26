"""
Updated 20.01.2023

@author: Tung Son Le
"""

try:
    import tkinter as tk
except ImportError:
    import Tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pyvisa
import re
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.backend_bases import key_press_handler
import pandas as pd
import time


def validate_p(new_value):
    if (new_value.replace(".", "", 1).isnumeric() and float(new_value) >= 0) or new_value == "":
        return True
    return False


def validate_n(new_value):
    try:
        if (new_value.replace(".", "", 1).isnumeric()) or new_value == "" or new_value == "-" or float(new_value):
            return True
    except ValueError:
        return False
    return False


def validate_file_location(file_name):
    invalid_chars = re.search(r'[/"*?<>|]', file_name)
    if invalid_chars:
        return False
    return True


def validate_file_name(file_name):
    invalid_chars = re.search(r'[\\/:"*?<>|]', file_name)
    if invalid_chars:
        return False
    return True


class E4990A_fs_sweep:
    def __init__(self):
        # self.rm = pyvisa.ResourceManager()
        self.E4990A = None
        self.VISA_address = "USB0::0x2A8D::0x5F01::MY54303049::0::INSTR"

        self.channel = 1
        self.trace_number = 1
        self.allocate_channels_list = ["D1", "D12", "D1_2", "D112", "D1_1_2", "D123", "D1_2_3", "D12_33", "D11_23",
                                       "D13_23", "D12_13", "D1234", "D1_2_3_4", "D12_34"]
        self.allocate_channels_value = "D1"
        self.num_of_traces_value = 4
        self.allocate_traces_list = ["D1", "D12", "D1_2", "D112", "D1_1_2", "D123", "D1_2_3", "D12_33", "D11_23",
                                     "D13_23", "D12_13", "D1234", "D1_2_3_4", "D12_34"]
        self.allocate_traces_value = "D12_34"
        self.trace_parameter_list = ["Z", "Y", "R", "X", "G", "B", "LS", "LP", "CS", "CP", "RS", "RP", "Q", "D", "TZ",
                                     "TY", "VAC", "IAC", "VDC", "IDC", "IMP", "ADM"]
        self.trace_parameter_list_y_axis = ["|Z|(Ohm)-data", "|Y|(S)-data", "R(Ohm)-data", "X(Ohm)-Data", "G(S)-Data",
                                            "B(S)-Data", "LS(H)-Data", "LP(H)-Data", "CS(F)-Data", "CP(F)-Data",
                                            "RS(Ohm)-Data", "RP(Ohm)-Data", "Q()-Data", "D()-Data", "theta-z(deg)-Data",
                                            "theta-y(deg)-Data", "Vac(V)-Data", "Iac(A)-DataC", "Vdc()-Data",
                                            "Idc()-Data", "Z-(S)-Data", "Y-(S)-Data"]
        self.trace_1_parameter = "Z"
        self.trace_2_parameter = "TZ"
        self.trace_3_parameter = "R"
        self.trace_4_parameter = "CS"

        self.frequency_start_value = 20.0
        self.frequency_stop_value = 20000000
        self.sampling_frequency_value = None
        self.number_of_point_value = 201
        self.ac_voltage_level_value = 0.5
        self.dc_bias_voltage_level_value = 0
        self.measurement_speed_value = 1

        self.continuous_value = 1
        self.target_traces_list = ["TRACe", "CHANnel", "DISPlayed"]
        self.pc_file_location_value = "C:\\Users\\cottbus_lab\\Desktop\\"
        self.pc_file_name_value = ""
        self.e4990a_file_location_value = "C:Users\\Instrument\\Desktop\\"
        self.e4990a_file_name_value = ""

        self.root = Tk()
        # configure the root window
        self.root.title("fs_sweep")
        self.root.geometry("1300x1000+0+0")
        # self.root.attributes('-fullscreen', False)

        # control frame
        self.side_frame = Frame(self.root)
        self.side_frame.pack(side=LEFT, fill=BOTH, expand=0)

        # plot frame
        self.plot_frame = Frame(self.root)
        self.plot_frame.pack(side=LEFT, fill=BOTH, expand=1)

        # Connection
        self.connection_frame = LabelFrame(self.side_frame, text="Connection")
        self.connection_frame.pack(fill=X, expand=0)

        self.connect_button = Button(self.connection_frame, text="Connect", width=12,
                                     command=self.connect_with_E4990A)
        self.connect_button.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.connection_indicator_canvas = Canvas(self.connection_frame, width=1, height=20)
        self.connection_indicator_canvas.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.connection_indicator = self.connection_indicator_canvas.create_oval(2, 2, 18, 18, fill="red")
        self.connection_status = Label(self.connection_frame, text="Offline")
        self.connection_status.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)

        # Settings
        self.settings_frame = LabelFrame(self.side_frame, text="Settings")
        self.settings_frame.pack(fill=X, expand=0)

        # Display
        self.display_frame = LabelFrame(self.settings_frame, labelanchor="n", text="Display")
        self.display_frame.pack(fill=X, expand=0)
        self.allocate_channels_frame = Frame(self.display_frame)
        self.allocate_channels_frame.pack(fill=X, expand=0)
        self.allocate_channels_label = Label(self.allocate_channels_frame, text="Allocate Channels", width=20, anchor=W)
        self.allocate_channels_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.allocate_channels_combobox = ttk.Combobox(self.allocate_channels_frame, width=17,
                                                       values=self.allocate_channels_list, state="readonly")
        self.allocate_channels_combobox.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.allocate_channels_combobox.set(self.allocate_channels_value)
        self.allocate_channels_combobox.bind('<<ComboboxSelected>>', lambda event: None)

        self.num_of_traces_value = IntVar()
        self.num_of_traces_value.set(4)
        self.num_of_traces_frame = Frame(self.display_frame)
        self.num_of_traces_frame.pack(fill=X, expand=0)
        self.num_of_traces_label = Label(self.num_of_traces_frame, text="Num Of Traces", width=20, anchor=W)
        self.num_of_traces_label.pack(side=LEFT, fill=X, expand=0, padx=2, pady=2)
        self.num_of_traces_1 = Radiobutton(self.num_of_traces_frame, text="1", variable=self.num_of_traces_value,
                                           value=1, indicator=0)
        self.num_of_traces_1.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.num_of_traces_2 = Radiobutton(self.num_of_traces_frame, text="2", variable=self.num_of_traces_value,
                                           value=2, indicator=0)
        self.num_of_traces_2.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.num_of_traces_3 = Radiobutton(self.num_of_traces_frame, text="3", variable=self.num_of_traces_value,
                                           value=3, indicator=0)
        self.num_of_traces_3.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.num_of_traces_4 = Radiobutton(self.num_of_traces_frame, text="4", variable=self.num_of_traces_value,
                                           value=4, indicator=0)
        self.num_of_traces_4.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)

        self.allocate_traces_frame = Frame(self.display_frame)
        self.allocate_traces_frame.pack(fill=X, expand=0)
        self.allocate_traces_label = Label(self.allocate_traces_frame, text="Allocate Traces", width=20, anchor=W)
        self.allocate_traces_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.allocate_traces_combobox = ttk.Combobox(self.allocate_traces_frame, width=17,
                                                     values=self.allocate_traces_list, state="readonly")
        self.allocate_traces_combobox.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.allocate_traces_combobox.set(self.allocate_traces_value)
        self.allocate_traces_combobox.bind('<<ComboboxSelected>>', lambda event: None)

        # Measurement
        self.measurement_frame = LabelFrame(self.settings_frame, labelanchor="n", text="Measurement")
        self.measurement_frame.pack(fill=X, expand=0)

        self.trace_1_frame = Frame(self.measurement_frame)
        self.trace_1_frame.pack(fill=X, expand=0)
        self.trace_1_label = Label(self.trace_1_frame, text="Trace 1", width=20, anchor=W)
        self.trace_1_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.trace_1_combobox = ttk.Combobox(self.trace_1_frame, width=17, values=self.trace_parameter_list,
                                             state="readonly")
        self.trace_1_combobox.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.trace_1_combobox.set(self.trace_1_parameter)
        self.trace_1_combobox.bind('<<ComboboxSelected>>', lambda event: None)

        self.trace_2_frame = Frame(self.measurement_frame)
        self.trace_2_frame.pack(fill=X, expand=0)
        self.trace_2_label = Label(self.trace_2_frame, text="Trace 2", width=20, anchor=W)
        self.trace_2_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.trace_2_combobox = ttk.Combobox(self.trace_2_frame, width=17, values=self.trace_parameter_list,
                                             state="readonly")
        self.trace_2_combobox.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.trace_2_combobox.set(self.trace_2_parameter)
        self.trace_2_combobox.bind('<<ComboboxSelected>>', lambda event: None)

        self.trace_3_frame = Frame(self.measurement_frame)
        self.trace_3_frame.pack(fill=X, expand=0)
        self.trace_3_label = Label(self.trace_3_frame, text="Trace 2", width=20, anchor=W)
        self.trace_3_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.trace_3_combobox = ttk.Combobox(self.trace_3_frame, width=17, values=self.trace_parameter_list,
                                             state="readonly")
        self.trace_3_combobox.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.trace_3_combobox.set(self.trace_3_parameter)
        self.trace_3_combobox.bind('<<ComboboxSelected>>', lambda event: None)

        self.trace_4_frame = Frame(self.measurement_frame)
        self.trace_4_frame.pack(fill=X, expand=0)
        self.trace_4_label = Label(self.trace_4_frame, text="Trace 2", width=20, anchor=W)
        self.trace_4_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.trace_4_combobox = ttk.Combobox(self.trace_4_frame, width=17, values=self.trace_parameter_list,
                                             state="readonly")
        self.trace_4_combobox.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.trace_4_combobox.set(self.trace_4_parameter)
        self.trace_4_combobox.bind('<<ComboboxSelected>>', lambda event: None)

        self.set_settings_frame = Frame(self.settings_frame)
        self.set_settings_frame.pack(fill=X, expand=0)
        self.set_settings_button = Button(self.settings_frame, text="Set Settings", command=self.set_settings)
        self.set_settings_button.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)

        # Input
        self.input_frame = LabelFrame(self.side_frame, text="Input")
        self.input_frame.pack(fill=X, expand=0)

        self.frequency_start_frame = Frame(self.input_frame)
        self.frequency_start_frame.pack(fill=X, expand=0)
        self.frequency_start_label = Label(self.frequency_start_frame, text="Frequency Start [Hz]", width=20, anchor=W)
        self.frequency_start_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.frequency_start_entry = Entry(self.frequency_start_frame, validate="key", width=20)
        self.frequency_start_entry.config(validatecommand=(self.frequency_start_entry.register(validate_p), '%P'))
        self.frequency_start_entry.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.frequency_start_entry.insert(END, str(self.frequency_start_value))
        self.frequency_start_entry.bind('<Return>', lambda event: None)

        self.frequency_stop_frame = Frame(self.input_frame)
        self.frequency_stop_frame.pack(fill=X, expand=0)
        self.frequency_stop_label = Label(self.frequency_stop_frame, text="Frequency Stop [Hz]", width=20, anchor=W)
        self.frequency_stop_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.frequency_stop_entry = Entry(self.frequency_stop_frame, validate="key", width=20)
        self.frequency_stop_entry.config(validatecommand=(self.frequency_stop_entry.register(validate_p), '%P'))
        self.frequency_stop_entry.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.frequency_stop_entry.insert(END, str(self.frequency_stop_value))
        self.frequency_stop_entry.bind('<Return>', lambda event: None)

        self.sampling_mode_value = IntVar()
        self.sampling_frame = Frame(self.input_frame)
        self.sampling_frame.pack(fill=X, expand=0)
        self.sampling_scale = Scale(self.sampling_frame, variable=self.sampling_mode_value, from_=0, to=1, length=50,
                                    sliderlength=25, orient=HORIZONTAL, showvalue=0, command=self.sampling_mode)
        self.sampling_scale.pack(side=LEFT, fill=X, expand=0, padx=2, pady=2)

        self.sampling_frequency_label = Label(self.sampling_frame, text="\u0394f", anchor=W)
        self.sampling_frequency_label.pack(side=LEFT, fill=X, expand=0, padx=2, pady=2)
        self.sampling_frequency_entry = Entry(self.sampling_frame, validate="key", width=7)
        self.sampling_frequency_entry.config(
            validatecommand=(self.sampling_frequency_entry.register(validate_p), '%P'))
        self.sampling_frequency_entry.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.sampling_frequency_entry.bind('<Return>', lambda event: None)

        self.number_of_point_label = Label(self.sampling_frame, text="Points (2-1601)", anchor=W)
        self.number_of_point_label.pack(side=LEFT, fill=X, expand=0, padx=2, pady=2)
        self.number_of_point_entry = Entry(self.sampling_frame, validate="key", width=1)
        self.number_of_point_entry.config(validatecommand=(self.number_of_point_entry.register(validate_p), '%P'))
        self.number_of_point_entry.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.number_of_point_entry.insert(END, str(self.number_of_point_value))
        self.number_of_point_entry.config(state=DISABLED)
        self.number_of_point_entry.bind('<Return>', lambda event: None)

        self.ac_voltage_level_frame = Frame(self.input_frame)
        self.ac_voltage_level_frame.pack(fill=X, expand=0)
        self.ac_voltage_level_label = Label(self.ac_voltage_level_frame, text="AC Voltage Level [V] (0-1V)", width=20,
                                            anchor=W)
        self.ac_voltage_level_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.ac_voltage_level_entry = Entry(self.ac_voltage_level_frame, validate="key", width=20)
        self.ac_voltage_level_entry.config(validatecommand=(self.ac_voltage_level_entry.register(validate_p), '%P'))
        self.ac_voltage_level_entry.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.ac_voltage_level_entry.insert(END, str(self.ac_voltage_level_value))
        self.ac_voltage_level_entry.bind('<Return>', lambda event: None)

        self.dc_bias_state_value = IntVar()
        self.dc_bias_state_value.set(0)
        self.dc_bias_frame = Frame(self.input_frame)
        self.dc_bias_frame.pack(fill=X, expand=0)
        self.dc_bias_state_off = Radiobutton(self.dc_bias_frame, text="OFF", variable=self.dc_bias_state_value,
                                             value=0, indicator=0)
        self.dc_bias_state_off.pack(side=LEFT, fill=X, expand=1, pady=2)
        self.dc_bias_state_on = Radiobutton(self.dc_bias_frame, text="ON", variable=self.dc_bias_state_value,
                                            value=1, indicator=0)
        self.dc_bias_state_on.pack(side=LEFT, fill=X, expand=1, pady=2)
        self.dc_bias_voltage_level_label = Label(self.dc_bias_frame, text="DC Bias(-40-40V)", width=0, anchor=W)
        self.dc_bias_voltage_level_label.pack(side=LEFT, fill=X, expand=1, pady=2)
        self.dc_bias_voltage_level_entry = Entry(self.dc_bias_frame, validate="key", width=20)
        self.dc_bias_voltage_level_entry.config(
            validatecommand=(self.dc_bias_voltage_level_entry.register(validate_n), '%P'))
        self.dc_bias_voltage_level_entry.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.dc_bias_voltage_level_entry.insert(END, str(self.dc_bias_voltage_level_value))
        self.dc_bias_voltage_level_entry.bind('<Return>', lambda event: None)

        self.measurement_speed_value = IntVar()
        self.measurement_speed_value.set(1)
        self.measurement_speed_frame = Frame(self.input_frame)
        self.measurement_speed_frame.pack(fill=X, expand=0)
        self.measurement_speed_label = Label(self.measurement_speed_frame, text="Measurement Speed", width=20, anchor=W)
        self.measurement_speed_label.pack(side=LEFT, fill=X, expand=0, padx=2, pady=2)
        self.measurement_speed_1 = Radiobutton(self.measurement_speed_frame, text="1",
                                               variable=self.measurement_speed_value, value=1, indicator=0)
        self.measurement_speed_1.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.measurement_speed_2 = Radiobutton(self.measurement_speed_frame, text="2",
                                               variable=self.measurement_speed_value, value=2, indicator=0)
        self.measurement_speed_2.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.measurement_speed_3 = Radiobutton(self.measurement_speed_frame, text="3",
                                               variable=self.measurement_speed_value, value=3, indicator=0)
        self.measurement_speed_3.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.measurement_speed_4 = Radiobutton(self.measurement_speed_frame, text="4",
                                               variable=self.measurement_speed_value, value=4, indicator=0)
        self.measurement_speed_4.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.measurement_speed_5 = Radiobutton(self.measurement_speed_frame, text="5",
                                               variable=self.measurement_speed_value, value=5, indicator=0)
        self.measurement_speed_5.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)

        self.set_input_frame = Frame(self.input_frame)
        self.set_input_frame.pack(fill=X, expand=0)
        self.set_input_button = Button(self.set_input_frame, text="Set Input",
                                       command=self.set_input)
        self.set_input_button.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)

        # Output
        self.output_frame = LabelFrame(self.side_frame, text="Output")
        self.output_frame.pack(fill=X, expand=0)

        self.dataformat_check_value = [IntVar() for _ in range(0, 2)]
        self.dataformat_frame = Frame(self.output_frame)
        self.dataformat_frame.pack(fill=X, expand=0)
        self.dataformat_label = Label(self.dataformat_frame, text="Dataformat", width=20, anchor=W)
        self.dataformat_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.csv_dataformat_checkbutton = Checkbutton(self.dataformat_frame, text="CSV",
                                                      variable=self.dataformat_check_value[0],
                                                      offvalue=False, onvalue=True, indicator=2)
        self.csv_dataformat_checkbutton.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.csv_dataformat_checkbutton.select()
        self.image_dataformat_checkbutton = Checkbutton(self.dataformat_frame, text="IMAGE",
                                                        variable=self.dataformat_check_value[1],
                                                        offvalue=False, onvalue=True, indicator=2)
        self.image_dataformat_checkbutton.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.image_dataformat_checkbutton.select()

        self.pc_check_value = IntVar()
        self.pc_file_frame = Frame(self.output_frame)
        self.pc_file_frame.pack(fill=X, expand=0)
        self.pc_checkbutton = Checkbutton(self.pc_file_frame, text="Save File in PC", variable=self.pc_check_value,
                                          offvalue=False, onvalue=True, indicator=2, command=self.pc_file)
        self.pc_checkbutton.pack(side=LEFT, fill=X, expand=0, padx=2, pady=2)
        self.pc_checkbutton.select()

        self.pc_file_location_frame = Frame(self.output_frame)
        self.pc_file_location_frame.pack(fill=X, expand=0)
        self.pc_file_location_label = Label(self.pc_file_location_frame, text="File Location", width=6, anchor=W)
        self.pc_file_location_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.pc_file_location_entry = Entry(self.pc_file_location_frame, validate="key", width=30)
        self.pc_file_location_entry.config(
            validatecommand=(self.pc_file_location_entry.register(validate_file_location), '%P'))
        self.pc_file_location_entry.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.pc_file_location_entry.insert(END, str(self.pc_file_location_value))
        self.pc_file_location_entry.bind('<Return>', lambda event: None)

        self.pc_file_name_frame = Frame(self.output_frame)
        self.pc_file_name_frame.pack(fill=X, expand=0)
        self.pc_file_name_label = Label(self.pc_file_name_frame, text="File Name", width=6, anchor=W)
        self.pc_file_name_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.pc_file_name_entry = Entry(self.pc_file_name_frame, validate="key", width=30)
        self.pc_file_name_entry.config(validatecommand=(self.pc_file_name_entry.register(validate_file_name), '%P'))
        self.pc_file_name_entry.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.pc_file_name_entry.bind('<Return>', lambda event: None)

        self.e4990a_check_value = IntVar()
        self.e4990a_file_frame = Frame(self.output_frame)
        self.e4990a_file_frame.pack(fill=X, expand=0)
        self.e4990a_checkbutton = Checkbutton(self.e4990a_file_frame, text="Save File in E4990A",
                                              variable=self.e4990a_check_value,
                                              offvalue=False, onvalue=True, indicator=2, command=self.e4990a_file)
        self.e4990a_checkbutton.pack(side=LEFT, fill=X, expand=0, padx=2, pady=2)

        self.e4990a_file_location_frame = Frame(self.output_frame)
        self.e4990a_file_location_frame.pack(fill=X, expand=0)
        self.e4990a_file_location_label = Label(self.e4990a_file_location_frame, text="File Location",
                                                width=6, anchor=W)
        self.e4990a_file_location_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.e4990a_file_location_entry = Entry(self.e4990a_file_location_frame, validate="key", width=30)
        self.e4990a_file_location_entry.config(validatecommand=(self.e4990a_file_location_entry.
                                                                register(validate_file_location), '%P'))
        self.e4990a_file_location_entry.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.e4990a_file_location_entry.insert(END, str(self.e4990a_file_location_value))
        self.e4990a_file_location_entry.bind('<Return>', lambda event: None)
        self.e4990a_file_location_entry.configure(state=DISABLED)

        self.e4990a_file_name_frame = Frame(self.output_frame)
        self.e4990a_file_name_frame.pack(fill=X, expand=0)
        self.e4990a_file_name_label = Label(self.e4990a_file_name_frame, text="File Name", width=6, anchor=W)
        self.e4990a_file_name_label.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.e4990a_file_name_entry = Entry(self.e4990a_file_name_frame, validate="key", width=30)
        self.e4990a_file_name_entry.config(validatecommand=(self.e4990a_file_name_entry.register(validate_file_name),
                                                            '%P'))
        self.e4990a_file_name_entry.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.e4990a_file_name_entry.bind('<Return>', lambda event: None)
        self.e4990a_file_name_entry.configure(state=DISABLED)

        self.run_frame = Frame(self.output_frame)
        self.run_frame.pack(fill=X, expand=0)
        self.run_button = Button(self.run_frame, text="Set Output", command=self.run)
        self.run_button.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)

        # Other
        self.other_frame = LabelFrame(self.side_frame, text="Other")
        self.other_frame.pack(fill=X, expand=0)

        self.plot_type = IntVar()
        self.plot_type.set(1)
        self.plot_data_frame = Frame(self.other_frame)
        self.plot_data_frame.pack(fill=X, expand=0)
        self.plot_data_from_pc_button = Button(self.plot_data_frame, text="Plot data", command=self.plot_data_open)
        self.plot_data_from_pc_button.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.plot = Frame(self.plot_data_frame)
        self.plot.pack(fill=X, expand=0)
        self.plot_pc_type = Radiobutton(self.plot, text="PC-Type", variable=self.plot_type, value=1)
        self.plot_pc_type.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)
        self.plot_e4990a_type = Radiobutton(self.plot, text="E4990A-Type", variable=self.plot_type, value=2)
        self.plot_e4990a_type.pack(side=LEFT, fill=X, expand=1, padx=2, pady=2)

        # Plot
        self.fig = plt.Figure()
        self.canvas = FigureCanvasTkAgg(self.fig, self.plot_frame)
        self.canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=1)
        self.toolbar = NavigationToolbar2Tk(self.canvas, self.plot_frame)
        self.toolbar.update()
        # self.fig.canvas.mpl_connect('key_press_event', key_press_handler)

        self.ax1 = self.fig.add_subplot(221)
        self.ax2 = self.fig.add_subplot(222)
        self.ax3 = self.fig.add_subplot(223)
        self.ax4 = self.fig.add_subplot(224)

        for ax in [self.ax1, self.ax2, self.ax3, self.ax4]:
            ax.set_xlabel("Frequency(Hz)")
            ax.grid()

        # self.connect_with_E4990A()

    def connect_with_E4990A(self):
        try:
            self.E4990A.close()
        except AttributeError:
            pass
        try:
            self.E4990A = self.rm.open_resource(self.VISA_address)
            self.connection_indicator_canvas.itemconfig(self.connection_indicator, fill="green")
            self.connection_status.config(text="Online")
        except pyvisa.errors.VisaIOError:
            self.connection_indicator_canvas.itemconfig(self.connection_indicator, fill="red")
            self.connection_status.config(text="Offline")
            print("Try to connect with E4990A again")

    def close(self):
        self.rm.close()

    def sampling_mode(self, sampling_mode_value):
        if sampling_mode_value == "0":
            self.sampling_frequency_entry.config(state=NORMAL)
            self.number_of_point_entry.config(state=DISABLED)
        else:
            self.sampling_frequency_entry.config(state=DISABLED)
            self.number_of_point_entry.config(state=NORMAL)

    def pc_file(self):
        if self.pc_check_value.get():
            self.pc_file_location_entry.configure(state=NORMAL)
            self.pc_file_name_entry.configure(state=NORMAL)
        else:
            self.pc_file_location_entry.configure(state=DISABLED)
            self.pc_file_name_entry.configure(state=DISABLED)

    def e4990a_file(self):
        if self.e4990a_check_value.get():
            self.e4990a_file_location_entry.configure(state=NORMAL)
            self.e4990a_file_name_entry.configure(state=NORMAL)
        else:
            self.e4990a_file_location_entry.configure(state=DISABLED)
            self.e4990a_file_name_entry.configure(state=DISABLED)

    def set_settings(self):
        self.allocate_channels_value = self.allocate_channels_combobox.get()
        self.allocate_traces_value = self.allocate_traces_combobox.get()
        self.trace_1_parameter = self.trace_1_combobox.get()
        self.trace_2_parameter = self.trace_2_combobox.get()
        self.trace_3_parameter = self.trace_3_combobox.get()
        self.trace_4_parameter = self.trace_4_combobox.get()

        try:
            # Set Display
            self.allocate_channels(self.allocate_channels_value)
            self.num_of_traces(self.channel, self.trace_number, self.num_of_traces_value.get())
            self.allocate_traces(self.channel, self.allocate_traces_value)
            # Set Measurement
            self.trace(self.channel, 1, self.trace_1_parameter)
            self.trace(self.channel, 2, self.trace_2_parameter)
            self.trace(self.channel, 3, self.trace_3_parameter)
            self.trace(self.channel, 4, self.trace_4_parameter)
        except (AttributeError, pyvisa.errors.VisaIOError, pyvisa.errors.InvalidSession):
            self.connect_with_E4990A()

    def set_input(self):
        try:
            self.frequency_start_value = float(self.frequency_start_entry.get())
            if self.frequency_start_value < 20:
                self.frequency_start_value = 20
            elif 20 <= self.frequency_start_value < 1000:
                self.frequency_start_value = round(self.frequency_start_value, 3)
            elif self.frequency_start_value < 10000:
                self.frequency_start_value = round(self.frequency_start_value, 2)
            elif 10000 <= self.frequency_start_value < 100000:
                self.frequency_start_value = round(self.frequency_start_value, 1)
            elif 100000 <= self.frequency_start_value < 20000000:
                self.frequency_start_value = round(self.frequency_start_value, 0)
            else:
                self.frequency_start_value = 20000000
        except ValueError:
            pass

        try:
            self.frequency_stop_value = float(self.frequency_stop_entry.get())
            if self.frequency_stop_value < 20:
                self.frequency_stop_value = 20
            elif 20 <= self.frequency_stop_value < 1000:
                self.frequency_stop_value = round(self.frequency_stop_value, 3)
            elif self.frequency_stop_value < 10000:
                self.frequency_stop_value = round(self.frequency_stop_value, 2)
            elif 10000 <= self.frequency_stop_value < 100000:
                self.frequency_stop_value = round(self.frequency_stop_value, 1)
            elif 100000 <= self.frequency_stop_value < 20000000:
                self.frequency_stop_value = round(self.frequency_stop_value, 0)
            else:
                self.frequency_stop_value = 20000000
        except ValueError:
            pass
        if self.frequency_start_value > self.frequency_stop_value:
            self.frequency_stop_value = self.frequency_start_value
        self.frequency_start_entry.delete(0, END)
        self.frequency_start_entry.insert(END, str(self.frequency_start_value))
        self.frequency_stop_entry.delete(0, END)
        self.frequency_stop_entry.insert(END, str(self.frequency_stop_value))

        if self.sampling_frequency_entry.config("state")[-1] == NORMAL:
            min_sampling_frequency_value = (self.frequency_stop_value - self.frequency_start_value) / 1601
            max_sampling_frequency_value = (self.frequency_stop_value - self.frequency_start_value) / 2
            try:
                self.number_of_point_value = round(float(self.number_of_point_entry.get()))
                self.number_of_point_entry.config(state=NORMAL)
                self.number_of_point_entry.delete(0, END)
                self.number_of_point_entry.insert(END, str(self.number_of_point_value))
                self.number_of_point_entry.config(state=DISABLED)
            except ValueError:
                self.number_of_point_entry.config(state=NORMAL)
                self.number_of_point_entry.delete(0, END)
                self.number_of_point_entry.insert(END, str(self.number_of_point_value))
                self.number_of_point_entry.config(state=DISABLED)
            if self.frequency_start_value != self.frequency_stop_value:
                try:
                    self.sampling_frequency_value = float(self.sampling_frequency_entry.get())
                    if self.sampling_frequency_value < min_sampling_frequency_value:
                        self.number_of_point_value = 1601
                        self.sampling_frequency_value = round(min_sampling_frequency_value, 3)
                        self.sampling_frequency_entry.delete(0, END)
                        self.sampling_frequency_entry.insert(END, str(self.sampling_frequency_value))
                    elif self.sampling_frequency_value > max_sampling_frequency_value:
                        self.number_of_point_value = 2
                        self.sampling_frequency_value = round(max_sampling_frequency_value, 3)
                        self.sampling_frequency_entry.delete(0, END)
                        self.sampling_frequency_entry.insert(END, str(self.sampling_frequency_value))
                    else:
                        self.number_of_point_value = int(round(
                            (self.frequency_stop_value - self.frequency_start_value) / self.sampling_frequency_value))
                        self.sampling_frequency_value = round((self.frequency_stop_value - self.frequency_start_value)
                                                              / self.number_of_point_value, 3)
                        self.sampling_frequency_entry.delete(0, END)
                        self.sampling_frequency_entry.insert(END, str(self.sampling_frequency_value))
                    self.number_of_point_entry.config(state=NORMAL)
                    self.number_of_point_entry.delete(0, END)
                    self.number_of_point_entry.insert(END, str(self.number_of_point_value))
                    self.number_of_point_entry.config(state=DISABLED)
                except ValueError:
                    self.sampling_frequency_value = (self.frequency_stop_value - self.frequency_start_value) \
                                                    / self.number_of_point_value
                    self.sampling_frequency_value = round(self.sampling_frequency_value, 3)
                    self.sampling_frequency_entry.delete(0, END)
                    self.sampling_frequency_entry.insert(END, str(self.sampling_frequency_value))
            else:
                self.sampling_frequency_entry.delete(0, END)
                self.sampling_scale.set(1)
                self.sampling_frequency_entry.config(state=DISABLED)
                self.number_of_point_entry.config(state=NORMAL)
                if self.number_of_point_value > 1601:
                    self.number_of_point_value = 1601
                    self.number_of_point_entry.delete(0, END)
                    self.number_of_point_entry.insert(END, str(self.number_of_point_value))
                elif self.number_of_point_value < 2:
                    self.number_of_point_value = 2
                    self.number_of_point_entry.delete(0, END)
                    self.number_of_point_entry.insert(END, str(self.number_of_point_value))
        else:
            try:
                self.number_of_point_value = round(float(self.number_of_point_entry.get()))
                self.number_of_point_entry.delete(0, END)
                self.number_of_point_entry.insert(END, str(self.number_of_point_value))
            except ValueError:
                self.number_of_point_entry.delete(0, END)
                self.number_of_point_entry.insert(END, str(self.number_of_point_value))
            if self.number_of_point_value > 1601:
                self.number_of_point_value = 1601
                self.number_of_point_entry.delete(0, END)
                self.number_of_point_entry.insert(END, str(self.number_of_point_value))
            elif self.number_of_point_value < 2:
                self.number_of_point_value = 2
                self.number_of_point_entry.delete(0, END)
                self.number_of_point_entry.insert(END, str(self.number_of_point_value))
            if self.frequency_start_value != self.frequency_stop_value:
                self.sampling_frequency_value = round((self.frequency_stop_value - self.frequency_start_value)
                                                      / self.number_of_point_value, 3)
                self.sampling_frequency_entry.config(state=NORMAL)
                self.sampling_frequency_entry.delete(0, END)
                self.sampling_frequency_entry.insert(END, str(self.sampling_frequency_value))
                self.sampling_frequency_entry.config(state=DISABLED)
            else:
                self.sampling_frequency_entry.config(state=NORMAL)
                self.sampling_frequency_entry.delete(0, END)
                self.sampling_frequency_entry.config(state=DISABLED)

        try:
            self.ac_voltage_level_value = float(self.ac_voltage_level_entry.get())
            if self.ac_voltage_level_value < 0:
                self.ac_voltage_level_value = 0
            elif self.ac_voltage_level_value > 1:
                self.ac_voltage_level_value = 1
            else:
                self.ac_voltage_level_value = round(self.ac_voltage_level_value, 3)
        except ValueError:
            pass
        self.ac_voltage_level_entry.delete(0, END)
        self.ac_voltage_level_entry.insert(END, str(self.ac_voltage_level_value))

        try:
            self.dc_bias_voltage_level_value = float(self.dc_bias_voltage_level_entry.get())
            if self.dc_bias_voltage_level_value < -40:
                self.dc_bias_voltage_level_value = -40
            elif self.dc_bias_voltage_level_value > 40:
                self.dc_bias_voltage_level_value = 40
            else:
                self.dc_bias_voltage_level_value = round(self.dc_bias_voltage_level_value, 3)
        except ValueError:
            pass
        self.dc_bias_voltage_level_entry.delete(0, END)
        self.dc_bias_voltage_level_entry.insert(END, str(self.dc_bias_voltage_level_value))

        try:
            self.frequency_start(self.channel, self.frequency_start_value)
            self.frequency_stop(self.channel, self.frequency_stop_value)
            self.number_of_points(self.channel, self.number_of_point_value)
            self.ac_voltage_level(self.channel, self.ac_voltage_level_value)
            self.dc_bias_state(self.channel, self.dc_bias_state_value.get())
            self.dc_bias_voltage_level(self.channel, self.dc_bias_voltage_level_value)
            self.measurement_speed(self.channel, self.measurement_speed_value.get())
        except (AttributeError, pyvisa.errors.VisaIOError, pyvisa.errors.InvalidSession):
            self.connect_with_E4990A()

    def run(self):
        self.num_of_traces(self.channel, self.trace_number, self.num_of_traces_value.get())  # set number of traces
        # again to save csv file correctly in PC
        try:
            self.continuous_initiation(self.channel, 1)
            self.continuous_initiation(self.channel, 0)
            while True:
                time.sleep(0.25)
                if self.value_of_OSCR() == 0:
                    break
            self.auto_scale(self.channel, 1)
            self.auto_scale(self.channel, 2)
            self.auto_scale(self.channel, 3)
            self.auto_scale(self.channel, 4)

            if self.pc_check_value.get():
                self.pc_file_location_value = self.pc_file_location_entry.get()
                self.pc_file_name_value = self.pc_file_name_entry.get()
                if self.dataformat_check_value[0].get() == 1:
                    header = ["Frequency(Hz)"]
                    data_list = [list(np.linspace(self.get_frequency_start(self.channel),
                                                  self.get_frequency_stop(self.channel),
                                                  int(self.get_number_of_points(self.channel))))]
                    print(data_list)
                    for i in range(self.num_of_traces_value.get()):
                        self.select_trace(self.channel, i + 1)
                        if self.get_trace_parameter(self.channel, i + 1) in ["IMP", "ADM"]:
                            header.append(f"{self.get_trace_parameter(self.channel, i + 1)}-Real(S)-data")
                            data_list.append(
                                [x for i, x in enumerate(self.get_trace_value(self.channel)) if i % 2 == 0])
                            header.append(f"{self.get_trace_parameter(self.channel, i + 1)}-Imag(S)-data")
                            data_list.append(
                                [x for i, x in enumerate(self.get_trace_value(self.channel)) if i % 2 != 0])
                        else:
                            index = self.trace_parameter_list.index(self.get_trace_parameter(self.channel, i + 1))
                            header.append(self.trace_parameter_list_y_axis[index])
                            data_list.append(
                                [x for i, x in enumerate(self.get_trace_value(self.channel)) if i % 2 == 0])
                    df = pd.DataFrame(data_list).transpose()
                    df.columns = header
                    try:
                        df.to_csv(self.pc_file_location_value + self.pc_file_name_value + ".csv", index=False)
                    except PermissionError:
                        messagebox.showerror(title="Error",
                                             message=f"File {self.pc_file_location_value}"
                                                     f"{self.pc_file_name_value}.csv is being opened.\n"
                                                     f"If you want to overwrite, please close this file and run again.")
                    self.plot_data_live(header, data_list)
                if self.dataformat_check_value[1].get() == 1:
                    self.fig.savefig(self.pc_file_location_value + self.pc_file_name_value + ".png")
            if self.e4990a_check_value.get():
                self.e4990a_file_location_value = self.e4990a_file_location_entry.get()
                self.e4990a_file_name_value = self.e4990a_file_name_entry.get()
                if self.dataformat_check_value[0].get() == 1:
                    self.save_csv((self.e4990a_file_location_value + self.e4990a_file_name_value + ".csv"),
                                  self.target_traces_list[1])
                if self.dataformat_check_value[1].get() == 1:
                    self.save_image((self.e4990a_file_location_value + self.e4990a_file_name_value + ".png"))
        except (AttributeError, pyvisa.errors.VisaIOError, pyvisa.errors.InvalidSession):
            self.connect_with_E4990A()

    def plot_data_live(self, header, data_list):
        sample = None
        count = 1
        if all(x == data_list[0][0] for x in data_list[0]):
            sample = list(range(0, len(data_list[0])))
        self.fig.suptitle("Live")
        for index, ax in enumerate([self.ax1, self.ax2, self.ax3, self.ax4]):
            ax.clear()
            try:
                if "Real" in header[count]:
                    ax.plot(data_list[count], data_list[count + 1])
                    ax.set_xlabel(header[count])
                    count += 1
                    ax.set_ylabel(header[count])
                    count += 1
                else:
                    ax.set_ylabel(header[count])
                    if sample is None:
                        ax.plot(data_list[0], data_list[count])
                        ax.set_xlabel(f"{header[0]}")
                    else:
                        ax.plot(sample, data_list[count])
                        ax.set_xlabel(f"Sample with {header[0]} = {data_list[0][0]}")
                    count += 1
            except IndexError:
                pass
            ax.grid()
        self.canvas.draw()

    def plot_data_open(self):
        count = 1
        filepath = filedialog.askopenfilename(filetypes=[("CSV files", ".csv"), ("All files", ".*")])
        try:
            if self.plot_type.get() == 1:
                data = pd.read_csv(filepath, header=0)
            else:
                data = pd.read_csv(filepath, skiprows=lambda
                    i: i in range(0, 4) or i == (sum(1 for _ in open(filepath)) - 1), header=0)
            header = data.columns
            data_list = []
            sample = None
            for index in range(0, len(header)):
                data_list.append([float(x) for x in data[data.columns[index]].tolist()])
            if all(x == data_list[0][0] for x in data_list[0]):
                sample = list(range(0, len(data_list[0])))
            self.fig.suptitle(filepath)
            for index, ax in enumerate([self.ax1, self.ax2, self.ax3, self.ax4]):
                ax.clear()
                try:
                    if "Real" in header[count]:
                        ax.plot(data_list[count], data_list[count + 1])
                        ax.set_xlabel(header[count])
                        count += 1
                        ax.set_ylabel(header[count])
                        count += 1
                    else:
                        ax.set_ylabel(header[count])
                        if sample is None:
                            ax.plot(data_list[0], data_list[count])
                            ax.set_xlabel(f"{header[0]}")
                        else:
                            ax.plot(sample, data_list[count])
                            ax.set_xlabel(f"Sample with {header[0]} = {data_list[0][0]}")
                        count += 1
                except IndexError:
                    pass
                ax.grid()
        except UnicodeError:
            for ax in [self.ax1, self.ax2, self.ax3, self.ax4]:
                self.fig.suptitle(filepath)
                ax.clear()
                ax.set_xlabel("Frequency(Hz)")
                ax.grid()
        self.canvas.draw()

    # Settings
    def allocate_channels(self, allocate_channels_value):
        self.E4990A.write(':DISPlay:SPLit %s' % allocate_channels_value)

    def num_of_traces(self, channel, trace, num_of_traces_value):
        self.E4990A.write(':CALCulate%d:PARameter%d:COUNt %d' % (channel, trace, num_of_traces_value))

    def allocate_traces(self, channel, allocate_traces_value):
        self.E4990A.write(':DISPlay:WINDow%d:SPLit %s' % (channel, allocate_traces_value))

    def trace(self, channel, trace, trace_parameter):
        self.E4990A.write(':CALCulate%d:PARameter%d:DEFine %s' % (channel, trace, trace_parameter))

    # Input
    def frequency_start(self, channel, frequency_start_value):
        self.E4990A.write(':SENSe%d:FREQuency:STAR %G' % (channel, frequency_start_value))

    def frequency_stop(self, channel, frequency_stop_value):
        self.E4990A.write(':SENSe%d:FREQuency:STOP %G' % (channel, frequency_stop_value))

    def number_of_points(self, channel, number_of_point_value):
        self.E4990A.write(':SENSe%d:SWEep:POINts %d' % (channel, number_of_point_value))

    def ac_voltage_level(self, channel, ac_voltage_level_value):
        self.E4990A.write(':SOURce%d:VOLTage:LEVel:IMMediate:AMPLitude %G' % (channel, ac_voltage_level_value))

    def dc_bias_state(self, channel, dc_bias_state_value):
        self.E4990A.write(':SOURce%d:BIAS:STATe %d' % (channel, dc_bias_state_value))

    def dc_bias_voltage_level(self, channel, dc_bias_voltage_level_value):
        self.E4990A.write(':SOURce%d:BIAS:VOLTage:LEVel:IMMediate:AMPLitude %G'
                          % (channel, dc_bias_voltage_level_value))

    def measurement_speed(self, channel, measurement_speed):
        self.E4990A.write(':SENSe%d:APERture:TIME %d' % (channel, measurement_speed))

    # Output
    def continuous_initiation(self, channel, continuous_value):
        self.E4990A.write(':INITiate%d:CONTinuous %d' % (channel, continuous_value))

    def value_of_OSCR(self):
        value = self.E4990A.query_ascii_values(':STATus:OPERation:CONDition?')
        return value[0]

    def get_frequency_start(self, channel):
        value = self.E4990A.query_ascii_values(':SENSe%d:FREQuency:STARt?' % channel)
        return value[0]

    def get_frequency_stop(self, channel):
        value = self.E4990A.query_ascii_values(':SENSe%d:FREQuency:STOP?' % channel)
        return value[0]

    def get_number_of_points(self, channel):
        value = self.E4990A.query_ascii_values(':SENSe%d:SWEep:POINts?' % channel)
        return value[0]

    def get_trace_parameter(self, channel, trace):
        value = self.E4990A.query(':CALCulate%d:PARameter%d:DEFine?' % (channel, trace))
        return value.strip()

    def select_trace(self, channel, trace):
        self.E4990A.write(':CALCulate%d:PARameter%d:SELect' % (channel, trace))

    def get_trace_value(self, channel):
        return self.E4990A.query_ascii_values(':CALCulate%d:SELected:DATA:FDATa?' % channel)

    def auto_scale(self, channel, trace):
        self.E4990A.write(':DISPlay:WINDow%d:TRACe%d:Y:SCALe:AUTO' % (channel, trace))

    def cover_image_to_BinaryType(self):
        return self.E4990A.query_binary_values(':HCOPy:SDUMp:DATA:IMMediate?', 'B', False)

    def save_csv(self, csv_file_name, target_traces):
        self.E4990A.write(':MMEMory:STORe:TRACe "%s",%s' % (csv_file_name, target_traces))

    def save_image(self, image_file_name):
        self.E4990A.write(':MMEMory:STORe:IMAGe "%s"' % image_file_name)


if __name__ == "__main__":
    main = E4990A_fs_sweep()
    main.root.mainloop()
