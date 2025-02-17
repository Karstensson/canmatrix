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

from __future__ import absolute_import, division, print_function

import typing
from builtins import *

import canmatrix


def get_frame_info(db, frame):
    # type: (canmatrix.CanMatrix, canmatrix.Frame) -> typing.List[str]
    ret_array = []  # type: typing.List[str]

    if db.type == canmatrix.matrix_class.CAN:
        # frame-id
        if frame.arbitration_id.extended:
            ret_array.append("%3Xxh" % frame.arbitration_id.id)
        else:
            ret_array.append("%3Xh" % frame.arbitration_id.id)
    elif db.type == canmatrix.matrix_class.FLEXRAY:
        ret_array.append("TODO")
    elif db.type == canmatrix.matrix_class.SOMEIP:
        ret_array.append("%3Xh" % frame.header_id)

    # frame-Name
    ret_array.append(frame.name)
    ret_array.append(frame.size)
    ret_array.append(frame.effective_cycle_time)

    # determine send-type
    if "GenMsgSendType" in db.frame_defines:
        ret_array.append(frame.attribute("GenMsgSendType", db=db))
        #if "GenMsgDelayTime" in db.frame_defines:
            #ret_array.append(frame.attribute("GenMsgDelayTime", db=db))
        #else:
         #   ret_array.append("")
    else:
        ret_array.append("")
        #ret_array.append("")
    return ret_array


def get_signal(db, frame, sig, motorola_bit_format):
    # type: (canmatrix.CanMatrix, canmatrix.Frame, canmatrix.Signal, str) -> typing.Tuple[typing.List, typing.List]
    front_array = []  # type: typing.List[typing.Union[str, float]]
    back_array = []
    if motorola_bit_format == "msb":
        start_bit = sig.get_startbit(bit_numbering=1)
    elif motorola_bit_format == "msbreverse":
        start_bit = sig.get_startbit()
    else:  # motorolaBitFormat == "lsb"
        start_bit = sig.get_startbit(bit_numbering=1, start_little=True)

    # start bit
    front_array.append(start_bit)
    #size
    front_array.append(sig.size) #Length
    # eval byteorder (little_endian: intel == True / motorola == 0)
    if sig.is_little_endian:
        front_array.append("i")
    else:
        front_array.append("m")


    # start-value of signal available
    if(sig.initial_value <= 0.00005):
        front_array.append(0)
    else:
        front_array.append(sig.initial_value)
    # signal name
    front_array.append(sig.name)
    # write comment and size of signal in sheet
    # eval comment:
    comment = sig.comment if sig.comment else ""
    front_array.append(comment)
    # eval multiplex-info
    if frame.is_complex_multiplexed:
        for signal in frame.signals:
            if signal.muxer_for_signal is not None:
                comment = "Mode {} = {}".format(sig.muxer_for_signal, sig.multiplex)
    else:
        if sig.multiplex == 'Multiplexor':
            comment = "Mode Signal: " + comment
        elif sig.multiplex is not None:
            comment = "Mode " + str(sig.multiplex) + ":" + comment



    # SNA-value of signal available
    # Disabled support for not available values for now, since it is not needed
    #HK#if "GenSigSNA" in db.signal_defines:
    #HK#    sna = sig.attribute("GenSigSNA", db=db)
    #HK#    if sna is not None:
    #HK#        sna = sna[1:-1]
    #HK#    front_array.append(sna)
    #HK# no SNA-value of signal available / just for correct style:
    #HKelse:
    #HK    front_array.append(" ")

    # is a unit defined for signal?
    back_array.append(sig.min)
    back_array.append(sig.max)
    back_array.append(sig.factor)
    back_array.append(sig.offset)
    back_array.append(sig.is_signed)
    back_array.append(sig.unit)
    return front_array, back_array
