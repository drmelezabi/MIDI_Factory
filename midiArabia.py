import pandas as pd
import math

from miditoolkit.midi import parser as mid_parser

from typing import Dict, List, Optional


class MidiArabia:
    def __init__(self, path_midi: str, detect_rhythms=True):
        self.__detect_rhythms = detect_rhythms
        self.midi_obj = mid_parser.MidiFile(path_midi)
        self.ticks_per_beat = round(self.midi_obj.ticks_per_beat)
        self.time_signature = self.midi_obj.time_signature_changes[0]
        self.numerator = self.time_signature.numerator
        self._denominator = self.time_signature.denominator
        self.tempo = round(self.midi_obj.tempo_changes[0].tempo)
        self.rest_note = None

        self.__rhythms = {
            "4.0": "w",
            "3.0": "hd",
            f"{round(4 / 3,2)}": "ht",
            "2.0": "h",
            f"{round(3 / 2,2)}": "qd",
            f"{round(2 / 3,2)}": "qt",
            f"{round(1 + (1/4) ,2)}": "q^s",
            f"1.0": "q",
            f"{round(1 / 3,2)}": "et",
            f"{round(3 / 4,2)}": "ed",
            f"{round(1 / 2,2)}": "e",
            f"{round(1 / 4,2)}": "s",
            f"{round(1 / 6,2)}": "st",
        }
        self.__rhythms_floats = [float(item) for item in list(self.__rhythms.keys())]
        self.theme_start = "theme_start"
        self.phrase_start = "phrase_start"

        self.bar_duration = self.numerator * self.ticks_per_beat

        # Create empty Series
        self.__empty_series = pd.Series()

        # DataFrame basics
        self.midi_df = pd.DataFrame(
            {
                "note": self.__empty_series,
                "interval": self.__empty_series,
                "note_start_time": self.__empty_series,
                "duration": self.__empty_series,
                "slur": self.__empty_series,
                "velocity": self.__empty_series,
            }
        )

        # add markers columns
        for column in self.markers_set():
            self.midi_df[column] = self.__empty_series

    def __filter_close_values(self, num, rel_tol=0.03, abs_tol=0.03):
        return [
            item
            for item in self.__rhythms_floats
            if math.isclose(item, num, rel_tol=rel_tol, abs_tol=abs_tol)
        ]

    def __rhythms_define(self, duration: int) -> str:
        if self.__detect_rhythms:
            for i in range(-2, 3):
                temp = round(((duration + i) / self.ticks_per_beat), 2)
                num = self.__filter_close_values(temp)
                close_number = num[0] if (num) else temp
                rhythm = self.__rhythms.get(f"{close_number}")
                if rhythm:
                    return rhythm
                return temp
        else:
            return duration

    def markers_set(self) -> set:
        markers_set = set()
        for marker in self.midi_obj.markers:
            marker_titles = marker.text.split(",")
            for m in marker_titles:
                markers_set.add(m)
        return markers_set

    def __markers_dict(self) -> Dict[int, List[str]]:
        markers = {}
        for marker in self.midi_obj.markers:
            marker_titles = marker.text.split(",")
            markers[round(marker.time)] = marker_titles
        return markers

    def __get_quarter_note(self) -> Dict[int, int]:
        pitch_bends_raw = self.midi_obj.instruments[0].pitch_bends

        pitch_bends_filtered = list(
            filter(lambda item: item.pitch != 0, pitch_bends_raw)
        )

        pitch_bends = {}

        for pitch_bend in pitch_bends_filtered:
            pitch_bends[round(pitch_bend.time)] = pitch_bend.pitch

        return pitch_bends

    def __new_entry(
        self,
        start_time: int,
        duration: int,
        def_column: List[str],
        markers: Optional[List[str]],
        note: Optional[int] = None,
        velocity: Optional[int] = None,
        slur: int = 0,
        last_note: Optional[int] = None,
        last_note_end_time: Optional[int] = 0,
    ) -> Dict[str, int]:
        row = {key: 0 for key in def_column}
        row["note"] = note if (note and velocity) else self.rest_note
        row["note_start_time"] = start_time
        row["duration"] = self.__rhythms_define(duration)
        row["velocity"] = velocity
        row["interval"] = None

        if markers:
            for marker in markers:
                row[marker] = 1

        if slur:
            row["slur"] = slur

        markers_set = self.markers_set()

        if note != None and velocity:
            if self.theme_start in markers_set or self.phrase_start in markers_set:
                # handle intervals
                if row.get(self.theme_start) == 0 and row.get(self.phrase_start) == 0:
                    last_note_int = last_note or 0
                    current_note = row["note"] or 0
                    row["interval"] = (
                        current_note - last_note_int
                        if (start_time - last_note_end_time < self.bar_duration)
                        else None
                    )
            else:
                last_note_int = last_note or 0
                current_note = note or 0
                row["interval"] = (
                    current_note - last_note_int
                    if (start_time - last_note_end_time < self.bar_duration)
                    else None
                )

        return row

    def midi_to_dataFrame(self) -> pd.DataFrame:
        notes = self.midi_obj.instruments[0].notes
        quarter_tone_list = self.__get_quarter_note()
        midi_dataFrame = self.midi_df
        markers = self.__markers_dict()
        columns = midi_dataFrame.columns

        last_note: None | int = None
        last_note_end_time = 0
        quarter_tone = False

        for note in notes:
            pitch = note.pitch
            start_time = round(note.start)
            end_time = round(note.end)
            pitch_bend = quarter_tone_list.get(start_time)
            quarter_tone = pitch_bend / 341 * 0.5 if (pitch_bend) else 0
            rest_exist = start_time - last_note_end_time > 0

            # handle rest if exist
            if rest_exist:
                rest_start_time = last_note_end_time
                next_note_start_time = start_time

                rest_starts_bar = math.ceil((rest_start_time + 1) / self.bar_duration)
                rest_ends_bar = math.ceil(
                    (next_note_start_time - 1) / self.bar_duration
                )

                existed_rest_bars = rest_ends_bar - rest_starts_bar

                for i in range(existed_rest_bars):
                    bar_ends = (rest_starts_bar + i) * self.bar_duration

                    # Add a new rest event to the DataFrame
                    midi_dataFrame.loc[len(midi_dataFrame)] = self.__new_entry(
                        start_time=rest_start_time,
                        duration=bar_ends - rest_start_time,
                        def_column=columns,
                        markers=markers.get(rest_start_time),
                    )

                    # update new rest start point
                    rest_start_time = bar_ends

                # existed_rest_less_than_bar
                if next_note_start_time - rest_start_time:
                    # Add a new rest event to the DataFrame
                    midi_dataFrame.loc[len(midi_dataFrame)] = self.__new_entry(
                        start_time=rest_start_time,
                        duration=start_time - rest_start_time,
                        def_column=columns,
                        markers=markers.get(rest_start_time),
                    )

            note_starts_bar = math.ceil((start_time + 1) / self.bar_duration)
            note_ends_bar = math.ceil((end_time - 1) / self.bar_duration)
            nested_slur_note = False

            existed_slur_note = note_ends_bar - note_starts_bar
            ## handle slur notes
            if existed_slur_note:
                for i in range(existed_slur_note):
                    slur = 0
                    if nested_slur_note:
                        slur = 1
                    else:
                        nested_slur_note = not nested_slur_note

                    bar_ends = (note_starts_bar + i) * self.bar_duration
                    current_note = pitch + quarter_tone

                    # Add a new note event to the DataFrame
                    midi_dataFrame.loc[len(midi_dataFrame)] = self.__new_entry(
                        note=current_note,
                        start_time=start_time,
                        duration=bar_ends - start_time,
                        velocity=note.velocity,
                        slur=slur,
                        def_column=columns,
                        last_note=last_note,
                        last_note_end_time=last_note_end_time,
                        markers=markers.get(start_time),
                    )

                    # update new note start point
                    last_note_end_time = bar_ends
                    start_time = bar_ends
                    last_note = current_note

                bar_ends = (note_starts_bar + i) * self.bar_duration
                current_note = pitch + quarter_tone

                ## there is nested slur note and it duration is less than bar duration
                if end_time % self.bar_duration:
                    # Add a new note event to the DataFrame
                    midi_dataFrame.loc[len(midi_dataFrame)] = self.__new_entry(
                        note=current_note,
                        start_time=start_time,
                        duration=end_time - start_time,
                        velocity=note.velocity,
                        slur=1,
                        def_column=columns,
                        last_note=last_note,
                        last_note_end_time=last_note_end_time,
                        markers=markers.get(start_time),
                    )
                # update new note start point
                last_note_end_time = end_time
                last_note = current_note
            ## handle none-slur notes
            else:
                bar_ends = (note_starts_bar + i) * self.bar_duration
                current_note = pitch + quarter_tone

                # Add a new note event to the DataFrame
                midi_dataFrame.loc[len(midi_dataFrame)] = self.__new_entry(
                    note=current_note,
                    start_time=start_time,
                    duration=end_time - start_time,
                    velocity=note.velocity,
                    def_column=columns,
                    last_note=last_note,
                    last_note_end_time=last_note_end_time,
                    markers=markers.get(start_time),
                )

                # update new note start point
                last_note_end_time = end_time
                last_note = current_note

        return midi_dataFrame
