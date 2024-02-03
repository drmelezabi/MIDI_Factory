## **MIDI Factory**
#### Import Libraries
import mido
import pandas as pd
import xlwings as xw


#### Create empty Series
empty_series = pd.Series()
### Create DataFrame with empty Series
midi_df = pd.DataFrame(
    {
        "note": empty_series,
        "note_start_time": empty_series,
        "duration": empty_series,
        "Velocity": empty_series,
        "heard_interval": empty_series,
        "theme_begin": empty_series,
        "theme_end": empty_series,
    }
)
### Load your MIDI file
mid = mido.MidiFile("2000.mid")

# Initialize a dictionary to store note start times and velocities
note_start_times = {}
note_velocities = {}
techniques = {}
bar_duration = 0
RestNote = 128
heard_interval_for_new_melody_begin = None
new_melody_beginning_marker = ("theme_start",)
new_phrase_beginning_marker = ("phrase_start",)


for i, track in enumerate(mid.tracks):

    total_time = 0  # Initialize total time for each track
    end_time = 0  # Initialize end time for rests
    last_note_end_time = 0
    quarter_tone = False
    lastNote = None
    previous_melody_last_note_loc = None

    for message in track:
        total_time += message.time  # Update total time for each message

        if message.type == "marker":
            marker_time = message.time
            marker_titles = message.text.split(",")

            for marker_title in marker_titles:

                if marker_time in techniques:
                    techniques[marker_time].append(marker_title)
                else:
                    techniques[marker_time] = [marker_title]

                if marker_title not in midi_df.columns:
                    if midi_df.shape[0]:
                        midi_df[marker_title] = 0
                    else:
                        midi_df[marker_title] = empty_series

        elif message.type == "pitchwheel":
            if message.pitch:
                quarter_tone = not quarter_tone

        elif message.type == "note_on" and message.velocity > 0:
            # Store the start time and velocity of the note
            note_start_times[message.note] = total_time
            note_velocities[message.note] = message.velocity

        elif message.type == "note_off":
            # Calculate the duration using the stored start time
            start_time = note_start_times.get(message.note, 0)
            duration = total_time - start_time
            velocity = note_velocities.get(message.note, 0)

            # Get current event's markers
            markers = []
            if techniques.get(start_time):
                markers = techniques.get(start_time)

            """
            # handel theme state
            if(new_melody_beginning_marker in markers or not lastNote):
                row_temp['theme_begin'] = 1
                row_temp['phrase_begin'] = 1
                row_temp['heard_interval'] = heard_interval_for_new_melody_begin
                if(previous_melody_last_note_loc):
                    midi_df.loc[previous_melody_last_note_loc]["theme_end"] = 1
                    midi_df.loc[previous_melody_last_note_loc]["phrase_end"] = 1
            else:
                row_temp['theme_begin'] = 0
                row_temp['heard_interval'] = (current_note - lastNote) * 4
                if(previous_melody_last_note_loc):
                    midi_df.loc[previous_melody_last_note_loc]["theme_end"] = 0
            
            # handel phrase state
            if(new_melody_beginning_marker in markers):
                row_temp['phrase_begin'] = 1
                if(previous_melody_last_note_loc):
                    midi_df.loc[previous_melody_last_note_loc]["phrase_end"] = 1
            else:
                row_temp['theme_begin'] = 0
                if(previous_melody_last_note_loc):
                    midi_df.loc[previous_melody_last_note_loc]["phrase_end"] = 0
            """

            # If there is rest before note
            if start_time - last_note_end_time > 0:
                rest_start_time = last_note_end_time
                rest_total_duration = start_time - last_note_end_time

                rest_starts_bar = rest_start_time // bar_duration
                rest_ends_bar = (start_time - 1) // bar_duration

                # Detect Note Pitch
                current_note = RestNote

                if rest_starts_bar != rest_ends_bar:
                    rest_row_temp = {key: 0 for key in midi_df.columns}
                    for i in range(rest_ends_bar - rest_starts_bar):
                        bar_ends = (rest_starts_bar + 1 + i) * bar_duration
                        rest_row_temp = {key: 0 for key in midi_df.columns}
                        rest_row_temp["note"] = current_note
                        rest_row_temp["note_start_time"] = rest_start_time
                        rest_row_temp["duration"] = bar_ends - rest_start_time

                        if techniques.get(start_time):
                            markers = techniques.get(start_time)

                        if markers and len(markers):
                            for rest_marker in markers:
                                rest_row_temp[rest_marker] = 1

                        # Add a new rest event to the DataFrame
                        midi_df.loc[len(midi_df)] = rest_row_temp
                        rest_start_time = bar_ends
                        last_note_end_time = bar_ends

                    if start_time - rest_start_time:
                        last_rest_duration = start_time - rest_start_time

                        rest_row_temp = {key: 0 for key in midi_df.columns}
                        rest_row_temp["note"] = current_note
                        rest_row_temp["note_start_time"] = rest_start_time
                        rest_row_temp["duration"] = last_rest_duration

                        # Add a new event to the DataFrame
                        midi_df.loc[len(midi_df)] = rest_row_temp
                        last_note_end_time = bar_ends
                else:
                    rest_row_temp = {key: 0 for key in midi_df.columns}
                    rest_row_temp["note"] = current_note
                    rest_row_temp["note_start_time"] = rest_start_time
                    rest_row_temp["duration"] = rest_total_duration

                    if techniques.get(start_time):
                        markers = techniques.get(start_time)

                    if markers and len(markers):
                        for rest_marker in markers:
                            rest_row_temp[rest_marker] = 1

                    # Add a new rest event to the DataFrame
                    midi_df.loc[len(midi_df)] = rest_row_temp
                    rest_start_time = bar_ends
                    last_note_end_time = bar_ends
                    # print(f'Rest, Start Time: {last_note_end_time}, Duration: {start_time - last_note_end_time}')

            if quarter_tone:
                current_note = message.note - 0.5
                quarter_tone = not quarter_tone
            else:
                current_note = message.note

            row_temp = {key: 0 for key in midi_df.columns}

            if markers and len(markers):
                for marker in markers:
                    row_temp[marker] = 1

            row_temp["note"] = current_note
            row_temp["note_start_time"] = start_time
            row_temp["duration"] = duration
            row_temp["Velocity"] = velocity
            # print(f'Note: {message.note}, Start_Time: {start_time}, Duration: {duration}, Velocity: {velocity}')

            # Add a new event to the DataFrame
            midi_df.loc[len(midi_df)] = row_temp
            # previous_melody_last_note_loc = len(midi_df)
            lastNote = current_note

            last_note_end_time = start_time + duration
        elif message.type == "time_signature":
            bar_duration = (
                message.numerator * message.denominator * message.clocks_per_click
            )


# Connect to the workbook
wb = xw.Book("test.xlsx")

# Select the worksheet
ws = wb.sheets["Sheet1"]

# Delete the existing table (if it exists)
if ws.tables:
    ws.range(ws.tables[0].range.address).api.Delete()

# Add a new table to the worksheet
table = ws.tables.add(source=ws.range("A1"), name="MyTable")

# Update the table with the new data
table.update(midi_df)

# Show the updated range of cells in a new window
table.range.api.Select()

# print(techniques)
