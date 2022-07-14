# -*- coding: utf-8 -*-
# Copyright (c) 2013, Eduard Broecker
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without modification, are permitted provided that
# the following conditions are met:
#
#    Redistributions of source code must retain the above copyright notice, this list of conditions and the
#    following disclaimer.
#    Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the
#    following disclaimer in the documentation and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED
# WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A
# PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY
# DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
# PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
# OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH
# DAMAGE.

#
# this script exports xls-files from a canmatrix-object
# xls-files are the can-matrix-definitions displayed in Excel

from __future__ import absolute_import, division, print_function

import decimal
import logging
import typing
from builtins import *

import past.builtins
import xlrd
import xlwt

import canmatrix
import canmatrix.formats.xls_common
import sys
logger = logging.getLogger(__name__)
default_float_factory = decimal.Decimal

# Font Size : 8pt * 20 = 160
# font = 'font: name Arial Narrow, height 160'
font = 'font: name Verdana, height 160'

if xlwt is not None:
    sty_header = xlwt.easyxf(font + ', bold on; align: vertical center, horizontal center',
                             'pattern: pattern solid, fore-colour rose')
    sty_norm = xlwt.easyxf(font + ', colour black')
    sty_first_frame = xlwt.easyxf(font + ', colour black; borders: top thin')
    sty_white = xlwt.easyxf(font + ', colour white')

    # ECU Matrix-Styles
    sty_green = xlwt.easyxf('pattern: pattern solid, fore-colour light_green')
    sty_green_first_frame = xlwt.easyxf('pattern: pattern solid, fore-colour light_green; borders: top thin')
    sty_sender = xlwt.easyxf('pattern: pattern 0x04, fore-colour gray25')
    sty_sender_first_frame = xlwt.easyxf('pattern: pattern 0x04, fore-colour gray25; borders: top thin')
    sty_sender_green = xlwt.easyxf('pattern: pattern 0x04, fore-colour gray25, back-colour light_green')
    sty_sender_green_first_frame = xlwt.easyxf(
        'pattern: pattern 0x04, fore-colour gray25, back-colour light_green; borders: top thin')


def write_ecu_matrix(ecus, sig, frame, worksheet, row, col, first_frame):
    # type: (typing.Sequence[str], typing.Optional[canmatrix.Signal], canmatrix.Frame, xlwt.Worksheet, int, int, xlwt.XFStyle) -> int
    # first-frame - style with borders:
    if first_frame == sty_first_frame:
        norm = sty_first_frame
        sender = sty_sender_first_frame
        norm_green = sty_green_first_frame
        sender_green = sty_sender_green_first_frame
    # consecutive-frame - style without borders:
    else:
        norm = sty_norm
        sender = sty_sender
        norm_green = sty_green
        sender_green = sty_sender_green

    # iterate over ECUs:
    for ecu_name in ecus:
        # every second ECU with other style
        if col % 2 == 0:
            loc_style = norm
            loc_style_sender = sender
        # every second ECU with other style
        else:
            loc_style = norm_green
            loc_style_sender = sender_green

        # write "s" "r" "r/s" if signal is sent, received or send and received by ECU
        if sig and ecu_name in sig.receivers and ecu_name in frame.transmitters:
            worksheet.write(row, col, label="r/s", style=loc_style_sender)
        elif sig and ecu_name in sig.receivers:
            worksheet.write(row, col, label="r", style=loc_style)
        elif ecu_name in frame.transmitters:
            worksheet.write(row, col, label="s", style=loc_style_sender)
        else:
            worksheet.write(row, col, label="", style=loc_style)
        col += 1
    # loop over ECUs ends here
    return col


def write_excel_line(worksheet, row, col, row_array, style):
    # type: (xlwt.Worksheet, int, int, typing.Sequence, xlwt.XFStyle) -> int
    for item in row_array:
        worksheet.write(row, col, label=item, style=style)
        col += 1
    return col


def dump(db, file, **options):
    # type: (canmatrix.CanMatrix, typing.IO, **typing.Any) -> None
    head_top = ['ID', 'Frame Name', 'Cycle Time [ms]', 'Launch Type', 'Launch Parameter', 'Signal Byte No.',
                'Signal Bit No.', 'Signal Name', 'Signal Function', 'Signal Length [Bit]', 'Signal Default',
                ' Signal Not Available', 'Byteorder']
    head_tail = ['Value',   'Name / Phys. Range', 'Function / Increment Unit']

    if len(options.get("additionalSignalAttributes", "")) > 0:
        additional_signal_columns = options.get("additionalSignalAttributes").split(",")  # type: typing.List[str]
    else:
        additional_signal_columns = []  # ["attributes['DisplayDecimalPlaces']"]

    if len(options.get("additionalFrameAttributes", "")) > 0:
        additional_frame_columns = options.get("additionalFrameAttributes").split(",")  # type: typing.List[str]
    else:
        additional_frame_columns = []  # ["attributes['DisplayDecimalPlaces']"]

    motorola_bit_format = options.get("xlsMotorolaBitFormat", "msbreverse")

    workbook = xlwt.Workbook(encoding='utf8')
#    ws_name = os.path.basename(filename).replace('.xls', '')
#    worksheet = workbook.add_sheet('K-Matrix ' + ws_name[0:22])
    worksheet = workbook.add_sheet('K-Matrix ')

    row_array = []  # type: typing.List[str]
    col = 0

    # write ECUs in first row:
    ecu_list = [ecu.name for ecu in db.ecus]

    row_array += head_top
    head_start = len(row_array)

    row_array += ecu_list
    for col in range(len(row_array)):
        worksheet.col(col).width = 1111
    tail_start = len(row_array)
    row_array += head_tail

    additional_frame_start = len(row_array)

    for col in range(tail_start, len(row_array)):
        worksheet.col(col).width = 3333

    for additionalCol in additional_frame_columns:
        row_array.append("frame." + additionalCol)
        col += 1

    for additionalCol in additional_signal_columns:
        row_array.append("signal." + additionalCol)
        col += 1

    write_excel_line(worksheet, 0, 0, row_array, sty_header)

    # set width of selected Cols
    worksheet.col(1).width = 5555
    worksheet.col(3).width = 3333
    worksheet.col(7).width = 5555
    worksheet.col(8).width = 7777
    worksheet.col(head_start).width = 1111
    worksheet.col(head_start + 1).width = 5555

    frame_hash = {}
    if db.type == canmatrix.matrix_class.CAN:
        logger.debug("Length of db.frames is %d", len(db.frames))
        for frame in db.frames:
            if frame.is_complex_multiplexed:
                logger.error("export complex multiplexers is not supported - ignoring frame %s", frame.name)
                continue
            frame_hash[int(frame.arbitration_id.id)] = frame
    else:
        frame_hash = {a.name:a for a in db.frames}


    # set row to first Frame (row = 0 is header)
    row = 1



    # iterate over the frames
    for idx in sorted(frame_hash.keys()):

        frame = frame_hash[idx]
        frame_style = sty_first_frame

        # sort signals:
        sig_hash = {"{:02d}{}".format(sig.get_startbit(), sig.name): sig for sig in frame.signals}

        # set style for first line with border
        sig_style = sty_first_frame

        additional_frame_info = [frame.attribute(frameInfo, default="") for frameInfo in additional_frame_columns]

        # iterate over signals
        row_array = []
        if len(sig_hash) == 0:  # Frames without signals
            row_array += canmatrix.formats.xls_common.get_frame_info(db, frame)
            for _ in range(5, head_start):
                row_array.append("")
            temp_col = write_excel_line(worksheet, row, 0, row_array, frame_style)
            temp_col = write_ecu_matrix(ecu_list, None, frame, worksheet, row, temp_col, frame_style)

            row_array = []
            for col in range(temp_col, additional_frame_start):
                row_array.append("")
            row_array += additional_frame_info
            for _ in additional_signal_columns:
                row_array.append("")
            write_excel_line(worksheet, row, temp_col, row_array, frame_style)
            row += 1
            continue

        # iterate over signals
        for sig_idx in sorted(sig_hash.keys()):
            sig = sig_hash[sig_idx]

            # if not first Signal in Frame, set style
            if sig_style != sty_first_frame:
                sig_style = sty_norm

            if sig.values.__len__() > 0:  # signals with value table
                val_style = sig_style
                # iterate over values in value table
                for val in sorted(sig.values.keys()):
                    row_array = canmatrix.formats.xls_common.get_frame_info(db, frame)
                    front_col = write_excel_line(worksheet, row, 0, row_array, frame_style)
                    if frame_style != sty_first_frame:
                        worksheet.row(row).level = 1

                    col = head_start
                    col = write_ecu_matrix(ecu_list, sig, frame, worksheet, row, col, frame_style)

                    # write Value
                    (frontRow, backRow) = canmatrix.formats.xls_common.get_signal(db, frame, sig, motorola_bit_format)
                    write_excel_line(worksheet, row, front_col, frontRow, sig_style)
                    backRow += additional_frame_info
                    for item in additional_signal_columns:
                        temp = getattr(sig, item, "")
                        backRow.append(temp)

                    write_excel_line(worksheet, row, col + 2, backRow, sig_style)
                    write_excel_line(worksheet, row, col, [val, sig.values[val]], val_style)

                    # no min/max here, because min/max has same col as values...
                    # next row
                    row += 1
                    # set style to normal - without border
                    sig_style = sty_white
                    frame_style = sty_white
                    val_style = sty_norm
                # loop over values ends here
            # no value table available
            else:
                row_array = canmatrix.formats.xls_common.get_frame_info(db, frame)
                front_col = write_excel_line(worksheet, row, 0, row_array, frame_style)
                if frame_style != sty_first_frame:
                    worksheet.row(row).level = 1

                col = head_start
                col = write_ecu_matrix(
                    ecu_list, sig, frame, worksheet, row, col, frame_style)
                (frontRow, backRow) = canmatrix.formats.xls_common.get_signal(db, frame, sig, motorola_bit_format)
                write_excel_line(worksheet, row, front_col, frontRow, sig_style)

                if float(sig.min) != 0 or float(sig.max) != 1.0:
                    backRow.insert(0, str("%g..%g" % (sig.min, sig.max)))  # type: ignore
                else:
                    backRow.insert(0, "")
                backRow.insert(0, "")

                backRow += additional_frame_info
                for item in additional_signal_columns:
                    temp = getattr(sig, item, "")
                    backRow.append(temp)

                write_excel_line(worksheet, row, col, backRow, sig_style)

                # next row
                row += 1
                # set style to normal - without border
                sig_style = sty_white
                frame_style = sty_white
        # loop over signals ends here
    # loop over frames ends here

    # frozen headings instead of split panes
    worksheet.set_panes_frozen(True)
    # in general, freeze after last heading row
    worksheet.set_horz_split_pos(1)
    worksheet.set_remove_splits(True)
    # save file
    workbook.save(file)


# ########################### load ###############################

def parse_value_table(value):
    # type: (str, str, int, typing.Callable) -> typing.Tuple
    value_table = dict()
    if len(value) > 0:
        tmp = value.split("\r\n")
        for val in tmp:
            data = val.split("=")
            value_table[int(data[0])] = data[1]
    return value_table


def read_additional_signal_attributes(signal, attribute_name, attribute_value):
    if not attribute_name.startswith("signal"):
        return
    if attribute_name.replace("signal.", "") in vars(signal):
        command_str = attribute_name + "="
        command_str += str(attribute_value)
        if len(str(attribute_value)) > 0:
            exec(command_str)
    else:
        pass


def load(file, **options):
    # type: (typing.IO, **typing.Any) -> canmatrix.CanMatrix
    motorola_bit_format = options.get("xlsMotorolaBitFormat", "msbreverse")
    float_factory = options.get("float_factory", default_float_factory)

    additional_inputs = dict()
    wb = xlrd.open_workbook(file_contents=file.read())
    sh = wb.sheet_by_index(0)
    db = canmatrix.CanMatrix()

    # Defines not imported...
    # db.add_ecu_defines("NWM-Stationsadresse", 'HEX 0 63')
    # db.add_ecu_defines("NWM-Knoten", 'ENUM  "nein","ja"')
    db.add_frame_defines("GenMsgDelayTime", 'INT 0 65535')
    db.add_frame_defines("GenMsgCycleTimeActive", 'INT 0 65535')
    db.add_frame_defines("GenMsgNrOfRepetitions", 'INT 0 65535')
    # db.addFrameDefines("GenMsgStartValue",  'STRING')
    launch_types = []  # type: typing.List[str]
    db.add_signal_defines("GenSigSNA", 'STRING')

    # eval search for correct columns:
    index = {}
    for i in range(sh.ncols):
        value = sh.cell(0, i).value
        if value == "ID":
            index['ID'] = i
        elif "Frame Name" in value:
            index['frameName'] = i
        elif "Cycle" in value:
            index['cycle'] = i
        elif "Tx Type" in value:
            index['tx_type'] = i
        elif "DLC" in value:
            index['dlc'] = i            
        elif "Start Bit" in value:
            index['startbit'] = i
        elif "Signal Name" in value:
            index['signalName'] = i
        elif "Signal Comment" in value:
            index['signalComment'] = i
        elif "Signal Length" in value:
            index['signalLength'] = i
        elif "Signal Init" in value:
            index['signalDefault'] = i
        elif "Min" in value:
            index['min'] = i      
        elif "Max" in value:
            index['max'] = i            
        elif "Factor" in value:
            index['factor'] = i
        elif "Offset" in value:
            index['offset'] = i
        elif "Unit" in value:
            index['unit'] = i                                                            
        elif "Byteorder" in value:
            index['byteorder'] = i
        elif "Signed" in value:
            index['signed'] = i                                   
        elif 'Value' in value:
                index['valueTable'] = i            

                

    index['ECUstart'] = index['signalComment'] + 1
    index['ECUend'] = index['valueTable']

    # ECUs:
    for x in range(index['ECUstart'], index['ECUend']):
        db.add_ecu(canmatrix.Ecu(sh.cell(0, x).value))

    # initialize:
    frame_id = None
    signal_name = ""
    new_frame = None

    for row_num in range(1, sh.nrows):
        # ignore empty row
        if len(sh.cell(row_num, index['ID']).value) == 0:
            break
        # new frame detected
        if sh.cell(row_num, index['ID']).value != frame_id:
            # new Frame
            frame_id = sh.cell(row_num, index['ID']).value
            frame_name = sh.cell(row_num, index['frameName']).value
            cycle_time = sh.cell(row_num, index['cycle']).value
            launch_type = sh.cell(row_num, index['tx_type']).value
            dlc = sh.cell(row_num, index['dlc']).value
            new_frame = canmatrix.Frame(frame_name, size=dlc)
            if frame_id.endswith("xh"):
                new_frame.arbitration_id = canmatrix.ArbitrationId(int(frame_id[:-2], 16), extended=True)
            else:
                new_frame.arbitration_id = canmatrix.ArbitrationId(int(frame_id[:-1], 16), extended=False)
            db.add_frame(new_frame)

            # eval launch_type
            if launch_type is not None:
                if len(launch_type) > 0:
                    new_frame.add_attribute("GenMsgSendType", launch_type)
                    if launch_type not in launch_types:
                        launch_types.append(launch_type)

            # eval cycle time
            try:
                cycle_time = int(cycle_time)
            except:
                cycle_time = 0
            new_frame.cycle_time = cycle_time

        # new signal detected
        if sh.cell(row_num, index['signalName']).value != signal_name \
                and len(sh.cell(row_num, index['signalName']).value) > 0:
            # new Signal
            receiver = []
            ##HK##start_byte = int(sh.cell(row_num, index['startbyte']).value)
            start_bit = int(sh.cell(row_num, index['startbit']).value)
            signal_name = sh.cell(row_num, index['signalName']).value.strip()
            signal_comment = sh.cell(
                row_num, index['signalComment']).value.strip()
            signal_length = int(sh.cell(row_num, index['signalLength']).value)
            signal_default = sh.cell(row_num, index['signalDefault']).value
            multiplex = None  # type: typing.Union[str, int, None]
            if signal_comment.startswith('Mode Signal:'):
                multiplex = 'Multiplexor'
                signal_comment = signal_comment[12:]
            elif signal_comment.startswith('Mode '):
                mux, signal_comment = signal_comment[4:].split(':', 1)
                multiplex = int(mux.strip())

            if index.get("byteorder", False):
                signal_byte_order = sh.cell(row_num, index['byteorder']).value

                if 'i' in signal_byte_order:
                    is_little_endian = True
                else:
                    is_little_endian = False
            else:
                is_little_endian = True  # Default Intel

            is_signed = False

            if signal_name != "-":
                for x in range(index['ECUstart'], index['ECUend']):
                    if 's' in sh.cell(row_num, x).value:
                        new_frame.add_transmitter(sh.cell(0, x).value.strip())
                    if 'r' in sh.cell(row_num, x).value:
                        receiver.append(sh.cell(0, x).value.strip())
                new_signal = canmatrix.Signal(
                    signal_name,
                    start_bit=start_bit,
                    size=int(signal_length),
                    is_little_endian=is_little_endian,
                    is_signed=is_signed,
                    receivers=receiver,
                    multiplex=multiplex)

                if not is_little_endian:
                    # motorola
                    if motorola_bit_format == "msb":
                        new_signal.set_startbit(start_bit, bitNumbering=1)
                    elif motorola_bit_format == "msbreverse":
                        new_signal.set_startbit(start_bit)
                    else:  # motorola_bit_format == "lsb"
                        new_signal.set_startbit(start_bit,
                            bitNumbering=1,
                            startLittle=True)

                new_frame.add_signal(new_signal)
                new_signal.add_comment(signal_comment)

        value_table = str(sh.cell(row_num, index['valueTable']).value)
        # .encode('utf-8')
        factor = sh.cell(row_num, index['factor']).value
        try:
            new_signal.factor = float_factory(factor)
        except:
            logger.warning(
                "Some error occurred while decoding scale of Signal %s: '%s'",
                signal_name, factor)
            new_signal.factor = 1
            
        new_signal.unit = str(sh.cell(row_num, index['unit']).value)
        new_signal.min = sh.cell(row_num, index['min']).value
        new_signal.max = sh.cell(row_num, index['max']).value
        new_signal.offset = float_factory(sh.cell(row_num, index['offset']).value)
        new_signal.initial_value = float_factory(signal_default)
        value_table = parse_value_table(value_table)

        if value_table is not None:
            for value, name in value_table.items():
                new_signal.add_values(value, name)

    for frame in db.frames:
        frame.update_receiver()
        frame.calc_dlc()

    launch_type_enum = "ENUM"
    launch_type_enum += ",".join([' "{}"'.format(launch_type) for launch_type in launch_types if launch_type])
    db.add_frame_defines("GenMsgSendType", launch_type_enum)

    db.set_fd_type()
    return db
