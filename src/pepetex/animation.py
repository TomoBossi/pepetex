import json
import argparse
import xml.etree.ElementTree as ET
from pathlib import Path

import utils
import errors
from namespaces import PREFIX_NAMESPACES

ANIMATION_ATTRIBS = {
    "by": {
        "name": "by",
        "type": int,
        "default": None,
        "validations": lambda by: True
    },
    "dir_barn": {
        "name": "dir",
        "type": str,
        "default": None,
        "validations": lambda dir: dir in ("inVertical", "inHorizontal", "outVertical", "outHorizontal")
    },
    "dir_blinds_randombar": {
        "name": "dir",
        "type": str,
        "default": None,
        "validations": lambda dir: dir in ("horizontal", "vertical")
    },
    "dir_box": {
        "name": "dir",
        "type": str,
        "default": None,
        "validations": lambda dir: dir in ("in", "out")
    },
    "dir_checkerboard": {
        "name": "dir",
        "type": str,
        "default": None,
        "validations": lambda dir: dir in ("across", "down")
    },
    "dir_strips": {
        "name": "dir",
        "type": str,
        "default": None,
        "validations": lambda dir: dir in ("downLeft", "upLeft", "downRight", "upRight")
    },
    "dir_wipe": {
        "name": "dir",
        "type": str,
        "default": None,
        "validations": lambda dir: dir in ("down", "left", "right", "up")
    },
    "dur": {
        "name": "dur",
        "type": int,
        "default": None,
        "validations": lambda dur: dur > 0
    },
    "spokes": {
        "name": "spokes",
        "type": int,
        "default": None,
        "validations": lambda spokes: spokes in (1, 2, 3, 4, 8)
    },
    "x": {
        "name": "x",
        "type": int,
        "default": None,
        "validations": lambda x: x > 0
    },
    "y": {
        "name": "y",
        "type": int,
        "default": None,
        "validations": lambda y: y > 0
    },
}
ANIMATION_SET_VISIBLE = """
<p:set>
    <p:cBhvr>
        <p:cTn dur="1" fill="hold">
            <p:stCondLst>
                <p:cond delay="0"/>
            </p:stCondLst>
        </p:cTn>
        <p:tgtEl>
            <p:spTgt spid="{uid}"/>
        </p:tgtEl>
        <p:attrNameLst>
            <p:attrName>style.visibility</p:attrName>
        </p:attrNameLst>
    </p:cBhvr>
    <p:to>
        <p:strVal val="visible"/>
    </p:to>
</p:set>
"""
ANIMATION_SET_HIDDEN = """
<p:set>
    <p:cBhvr>
        <p:cTn dur="1" fill="hold">
            <p:stCondLst>
                <p:cond delay="0"/>
            </p:stCondLst>
        </p:cTn>
        <p:tgtEl>
            <p:spTgt spid="{uid}"/>
        </p:tgtEl>
        <p:attrNameLst>
            <p:attrName>style.visibility</p:attrName>
        </p:attrNameLst>
    </p:cBhvr>
    <p:to>
        <p:strVal val="hidden"/>
    </p:to>
</p:set>
"""
ANIMATION_SET_HIDDEN_DELAY = """
<p:set>
    <p:cBhvr>
        <p:cTn dur="1" fill="hold">
            <p:stCondLst>
                <p:cond delay="{dur}"/>
            </p:stCondLst>
        </p:cTn>
        <p:tgtEl>
            <p:spTgt spid="{uid}"/>
        </p:tgtEl>
        <p:attrNameLst>
            <p:attrName>style.visibility</p:attrName>
        </p:attrNameLst>
    </p:cBhvr>
    <p:to>
        <p:strVal val="hidden"/>
    </p:to>
</p:set>
"""
ANIMATION_INTERPOLATE = """
<p:anim calcmode="lin" valueType="num">
    <p:cBhvr>
        <p:cTn dur="{{dur}}"/>
        <p:tgtEl>
            <p:spTgt spid="{{uid}}"/>
        </p:tgtEl>
        <p:attrNameLst>
            <p:attrName>{attrib}</p:attrName>
        </p:attrNameLst>
    </p:cBhvr>
    <p:tavLst>
        <p:tav tm="0">
            <p:val>
                <p:strVal val="{start}"/>
            </p:val>
        </p:tav>
        <p:tav tm="100000">
            <p:val>
                <p:strVal val="{end}"/>
            </p:val>
        </p:tav>
    </p:tavLst>
</p:anim>
"""
ANIMATION_INTERPOLATE_ADDITIVE = """
<p:anim calcmode="lin" valueType="num">
    <p:cBhvr additive="base">
        <p:cTn dur="{{dur}}" fill="hold"/>
        <p:tgtEl>
            <p:spTgt spid="{{uid}}"/>
        </p:tgtEl>
        <p:attrNameLst>
            <p:attrName>{attrib}</p:attrName>
        </p:attrNameLst>
    </p:cBhvr>
    <p:tavLst>
        <p:tav tm="0">
            <p:val>
                <p:strVal val="{start}"/>
            </p:val>
        </p:tav>
        <p:tav tm="100000">
            <p:val>
                <p:strVal val="{end}"/>
            </p:val>
        </p:tav>
    </p:tavLst>
</p:anim>
"""
ANIMATION_EFFECT = """
<p:animEffect transition="{transition}" filter="{filter}">
    <p:cBhvr>
        <p:cTn dur="{{dur}}"/>
        <p:tgtEl>
            <p:spTgt spid="{{uid}}"/>
        </p:tgtEl>
    </p:cBhvr>
</p:animEffect>
"""
ANIMATION_ROTATION = """
<p:animRot by="{by}">
    <p:cBhvr>
        <p:cTn dur="{dur}" fill="hold"/>
        <p:tgtEl>
            <p:spTgt spid="{uid}"/>
        </p:tgtEl>
        <p:attrNameLst>
            <p:attrName>r</p:attrName>
        </p:attrNameLst>
    </p:cBhvr>
</p:animRot>
"""
ANIMATION_SCALE = """
<p:animScale>
    <p:cBhvr>
        <p:cTn dur="{dur}" fill="hold"/>
        <p:tgtEl>
            <p:spTgt spid="{uid}"/>
        </p:tgtEl>
    </p:cBhvr>
    <p:by x="{x}" y="{y}"/>
</p:animScale>
"""
ANIMATIONS = {
    "appear": {
        "xml": ANIMATION_SET_VISIBLE,
        "presetID": 1,
        "presetClass": "entr", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": []
    },
    "basiczoom_in": {
        "xml":
            ANIMATION_SET_VISIBLE +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_w", "start": "0", "end": "ppt_w"}) +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_h", "start": "0", "end": "ppt_h"}),
        "presetID": 23,
        "presetClass": "entr", 
        "presetSubtype": 16,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "basiczoom_out": {
        "xml": 
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_w", "start": "ppt_w", "end": "0"}) +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_h", "start": "ppt_h", "end": "0"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 23,
        "presetClass": "exit", 
        "presetSubtype": 32,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "blinds_in": {
        "xml": 
            ANIMATION_SET_VISIBLE + 
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "blinds({dir})"}),
        "presetID": 3,
        "presetClass": "entr", 
        "presetSubtype": 10,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_blinds_randombar"], "horizontal")
        ]
    },
    "blinds_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "blinds({dir})"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 3,
        "presetClass": "exit", 
        "presetSubtype": 10,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_blinds_randombar"], "horizontal")
        ]
    },
    "box_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "box({dir})"}),
        "presetID": 4,
        "presetClass": "entr", 
        "presetSubtype": 16,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 2000),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_box"], "in")
        ]
    },
    "box_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "box({dir})"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 4,
        "presetClass": "exit", 
        "presetSubtype": 32,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 2000),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_box"], "out")
        ]
    },
    "checkerboard_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "checkerboard({dir})"}),
        "presetID": 5,
        "presetClass": "entr", 
        "presetSubtype": 10,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_checkerboard"], "across")
        ]
    },
    "checkerboard_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "checkerboard({dir})"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 5,
        "presetClass": "exit", 
        "presetSubtype": 10,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_checkerboard"], "across")
        ]
    },
    "circle_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "circle"}),
        "presetID": 0,
        "presetClass": "entr", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "circle_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "circle"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 0,
        "presetClass": "exit", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "diamond_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "diamond"}),
        "presetID": 0,
        "presetClass": "entr", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "diamond_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "diamond"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 0,
        "presetClass": "exit", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "disappear": {
        "xml": ANIMATION_SET_HIDDEN,
        "presetID": 1,
        "presetClass": "exit", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": []
    },
    "expand_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_w", "start": "#ppt_w*0.70", "end": "#ppt_w"}) +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_h", "start": "#ppt_h", "end": "#ppt_h"}) +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "fade"}),
        "presetID": 55,
        "presetClass": "entr", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 1000)
        ]
    },
    "fade_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "fade"}),
        "presetID": 10,
        "presetClass": "entr", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "fade_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "fade"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 10,
        "presetClass": "exit", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "fly_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_INTERPOLATE_ADDITIVE.format(**{"attrib": "ppt_x", "start": "#ppt_x", "end": "#ppt_x"}) +
            ANIMATION_INTERPOLATE_ADDITIVE.format(**{"attrib": "ppt_y", "start": "1+#ppt_h/2", "end": "#ppt_y"}),
        "presetID": 2,
        "presetClass": "entr", 
        "presetSubtype": 4,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "fly_out": {
        "xml": 
            ANIMATION_INTERPOLATE_ADDITIVE.format(**{"attrib": "ppt_x", "start": "#ppt_x", "end": "#ppt_x"}) +
            ANIMATION_INTERPOLATE_ADDITIVE.format(**{"attrib": "ppt_y", "start": "1+#ppt_h/2", "end": "#ppt_y"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 2,
        "presetClass": "entr", 
        "presetSubtype": 4,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "growturn_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_w", "start": "0", "end": "#ppt_w"}) +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_h", "start": "0", "end": "#ppt_h"}) +
            ANIMATION_INTERPOLATE.format(**{"attrib": "style.rotation", "start": "90", "end": "0"}) +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "fade"}),
        "presetID": 31,
        "presetClass": "entr", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 1000)
        ]
    },
    "peek_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_INTERPOLATE_ADDITIVE.format(**{"attrib": "ppt_y", "start": "#ppt_y+#ppt_h*1.125000", "end": "#ppt_y"}) +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "wipe({dir})"}),
        "presetID": 12,
        "presetClass": "entr", 
        "presetSubtype": 4,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_wipe"], "up")
        ]
    },
    "peek_out": {
        "xml": 
            ANIMATION_INTERPOLATE_ADDITIVE.format(**{"attrib": "ppt_y", "start": "#ppt_y", "end": "#ppt_y+#ppt_h*1.125000"}) +
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "wipe({dir})"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 12,
        "presetClass": "exit", 
        "presetSubtype": 4,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_wipe"], "down")
        ]
    },
    "randombars_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "randombar({dir})"}),
        "presetID": 14,
        "presetClass": "entr", 
        "presetSubtype": 10,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_blinds_randombar"], "horizontal")
        ]
    },
    "randombars_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "randombar({dir})"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 14,
        "presetClass": "exit", 
        "presetSubtype": 10,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_blinds_randombar"], "horizontal")
        ]
    },
    "scale": {
        "xml": ANIMATION_SCALE,
        "presetID": 8,
        "presetClass": "emph", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 2000),
            utils.set_attrib_default(ANIMATION_ATTRIBS["x"], 150000),
            utils.set_attrib_default(ANIMATION_ATTRIBS["y"], 150000)
        ]
    },
    "spin": {
        "xml": ANIMATION_ROTATION,
        "presetID": 8,
        "presetClass": "emph", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 2000),
            utils.set_attrib_default(ANIMATION_ATTRIBS["by"], 360*60000)
        ]
    },
    "split_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "barn({dir})"}),
        "presetID": 16,
        "presetClass": "entr", 
        "presetSubtype": 21,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_barn"], "inVertical")
        ]
    },
    "split_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "barn({dir})"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 16,
        "presetClass": "exit", 
        "presetSubtype": 21,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_barn"], "inVertical")
        ]
    },
    "stretch_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_w", "start": "0", "end": "#ppt_w"}) +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_h", "start": "#ppt_h", "end": "#ppt_h"}),
        "presetID": 17,
        "presetClass": "entr", 
        "presetSubtype": 10,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "stretch_out": {
        "xml":
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_w", "start": "ppt_w", "end": "0"}) +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_h", "start": "ppt_h", "end": "ppt_h"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 17,
        "presetClass": "exit", 
        "presetSubtype": 10,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "strips_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "strips({dir})"}),
        "presetID": 18,
        "presetClass": "entr", 
        "presetSubtype": 12,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_strips"], "downLeft")
        ]
    },
    "strips_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "strips({dir})"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 18,
        "presetClass": "exit", 
        "presetSubtype": 12,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_strips"], "downLeft")
        ]
    },
    "wheel_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "wheel({spokes})"}),
        "presetID": 0,
        "presetClass": "entr", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["spokes"], 4)
        ]
    },
    "wheel_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "wheel({spokes})"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 0,
        "presetClass": "exit", 
        "presetSubtype": 0,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["spokes"], 4)
        ]
    },
    "wipe_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "wipe({dir})"}),
        "presetID": 22,
        "presetClass": "entr", 
        "presetSubtype": 4,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_wipe"], "down")
        ]
    },
    "wipe_out": {
        "xml": 
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "wipe({dir})"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 22,
        "presetClass": "exit", 
        "presetSubtype": 4,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500),
            utils.set_attrib_default(ANIMATION_ATTRIBS["dir_wipe"], "down")
        ]
    },
    "zoom_in": {
        "xml": 
            ANIMATION_SET_VISIBLE +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_w", "start": "0", "end": "ppt_w"}) +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_h", "start": "0", "end": "ppt_h"}) +
            ANIMATION_EFFECT.format(**{"transition": "in", "filter": "fade"}),
        "presetID": 53,
        "presetClass": "entr", 
        "presetSubtype": 16,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    },
    "zoom_out": {
        "xml": 
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_w", "start": "ppt_w", "end": "0"}) +
            ANIMATION_INTERPOLATE.format(**{"attrib": "ppt_h", "start": "ppt_h", "end": "0"}) +
            ANIMATION_EFFECT.format(**{"transition": "out", "filter": "fade"}) +
            ANIMATION_SET_HIDDEN_DELAY,
        "presetID": 53,
        "presetClass": "exit", 
        "presetSubtype": 32,
        "fill": "hold",
        "attribs": [
            utils.set_attrib_default(ANIMATION_ATTRIBS["dur"], 500)
        ]
    }
}
TIMING_XML = """<?xml version='1.0' encoding='UTF-8'?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:timing>
        <p:tnLst>
            <p:par>
                <p:cTn dur="indefinite" nodeType="tmRoot" restart="never">
                    <p:childTnLst>
                        <p:seq concurrent="1" nextAc="seek">
                            <p:cTn dur="indefinite" nodeType="mainSeq">
                                <p:childTnLst>
                                <!--Animations-->
                                </p:childTnLst>
                            </p:cTn>
                            <p:prevCondLst>
                                <p:cond evt="onPrev">
                                    <p:tgtEl>
                                        <p:sldTgt />
                                    </p:tgtEl>
                                </p:cond>
                            </p:prevCondLst>
                            <p:nextCondLst>
                                <p:cond evt="onNext">
                                    <p:tgtEl>
                                        <p:sldTgt />
                                    </p:tgtEl>
                                </p:cond>
                            </p:nextCondLst>
                        </p:seq>
                    </p:childTnLst>
                </p:cTn>
            </p:par>
        </p:tnLst>
    </p:timing>
</p:sld>
"""
ANIMATION_AFTER_WITH_PREVIOUS_XML = """<?xml version='1.0' encoding='UTF-8'?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:par>
        <p:cTn fill="hold">
            <p:stCondLst>
                <p:cond delay="indefinite" />
                <p:cond evt="onBegin" delay="0">
                    <p:tn val="2" />
                </p:cond>
            </p:stCondLst>
            <p:childTnLst>
                <p:par>
                    <p:cTn fill="hold">
                        <p:stCondLst>
                            <p:cond delay="0" />
                        </p:stCondLst>
                        <p:childTnLst>
                        <!--Animation (withEffect or afterEffect)-->
                        </p:childTnLst>
                    </p:cTn>
                </p:par>
            </p:childTnLst>
        </p:cTn>
    </p:par>    
</p:sld>
"""
ANIMATION_AFTER_CLICK_XML = """<?xml version='1.0' encoding='UTF-8'?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:par>
        <p:cTn fill="hold">
            <p:stCondLst>
                <p:cond delay="indefinite"/>
            </p:stCondLst>
            <p:childTnLst>
                <p:par>
                    <p:cTn fill="hold">
                        <p:stCondLst>
                            <p:cond delay="0"/>
                        </p:stCondLst>
                        <p:childTnLst>
                        <!--Animation (clickEffect)-->
                        </p:childTnLst>
                    </p:cTn>
                </p:par>
            </p:childTnLst>
        </p:cTn>
    </p:par>    
</p:sld>
"""
ANIMATION_XML = """<?xml version='1.0' encoding='UTF-8'?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:par>
        <p:cTn fill="hold" nodeType="{nodeType}" presetClass="{presetClass}" presetID="{presetID}" presetSubtype="{presetSubtype}">
            <p:stCondLst>
                <p:cond delay="{delay}" />
            </p:stCondLst>
            <p:childTnLst>
            {animation}
            </p:childTnLst>
        </p:cTn>
    </p:par>    
</p:sld>
"""
def get_timing_index(slide: ET.Element) -> int:
    """
    Returns the <p:sld> children index at which the timing element 
    must be inserted in order to be compliant with PresentationML.
    See lines 1361-1367 at the end of page 3970 of ISO/IEC 29500-1 Third edition 2012-09-01.
    """
    i = 0
    for i, child in enumerate(slide):
        if utils.get_parsed_tag(child)["tag"] == "extLst":
            return i
    return i + 1

def get_timing_element(slide: ET.Element) -> ET.Element:
    """
    Returns the <p:timing> element of the slide. If there is no
    timing element it is added to the slide and returned. 
    """
    timing_elements = slide.findall(f"./{{{PREFIX_NAMESPACES['p']}}}timing")
    if timing_elements:
        return timing_elements[0]
    timing_element = ET.fromstring(TIMING_XML)
    timing_index = get_timing_index(slide)
    utils.insert_child_nodes(slide, timing_element, ".", timing_index)
    return timing_element

def build_animation_xml(
        uid: int,
        animation_name: str,
        animation_start: str = "click",
        animation_delay: int | None = None,
        animation_attribs: dict | None = None
    ) -> str:
    if animation_attribs is None:
        animation_attribs = {}
    animation_definition = ANIMATIONS[animation_name]
    for attrib in animation_definition["attribs"]:
        if attrib["name"] not in animation_attribs:
            animation_attribs[attrib["name"]] = attrib["default"]
    animation_attribs["uid"] = uid
    configured_animation = animation_definition["xml"].format(**animation_attribs)
    return ANIMATION_XML.format(**{
        "nodeType": {"click": "clickEffect", "after": "afterEffect", "with": "withEffect"}[animation_start],
        "presetClass": animation_definition["presetClass"],
        "presetID": animation_definition["presetID"],
        "presetSubtype": animation_definition["presetSubtype"],
        "delay": animation_delay or 0,
        "animation": configured_animation
    })

def extend_timing_element(
        timing_element: ET.Element,
        uid: int,
        animation_name: str,
        animation_start: str = "click",
        animation_index: int | None = None,
        animation_delay: int | None = None,
        animation_attribs: dict | None = None
    ) -> None:
    on_click = animation_start == "click"
    animation_element = ET.fromstring(build_animation_xml(uid, animation_name, animation_start, animation_delay, animation_attribs))
    animation_groups = timing_element.findall(f".//{{{PREFIX_NAMESPACES['p']}}}seq/{{{PREFIX_NAMESPACES['p']}}}cTn/{{{PREFIX_NAMESPACES['p']}}}childTnLst")[0]
    animation_wrapper = ET.fromstring(ANIMATION_AFTER_CLICK_XML)
    if len(animation_groups) == 0:
        if not on_click:
            animation_wrapper = ET.fromstring(ANIMATION_AFTER_WITH_PREVIOUS_XML)
        utils.append_child_nodes(animation_wrapper, animation_element, f".//{{{PREFIX_NAMESPACES['p']}}}par/{{{PREFIX_NAMESPACES['p']}}}cTn/{{{PREFIX_NAMESPACES['p']}}}childTnLst/{{{PREFIX_NAMESPACES['p']}}}par/{{{PREFIX_NAMESPACES['p']}}}cTn/{{{PREFIX_NAMESPACES['p']}}}childTnLst")
        utils.append_child_nodes(timing_element, animation_wrapper, f".//{{{PREFIX_NAMESPACES['p']}}}seq/{{{PREFIX_NAMESPACES['p']}}}cTn/{{{PREFIX_NAMESPACES['p']}}}childTnLst")
    else:
        pass
    print(ET.tostring(timing_element).decode("utf-8"))
    print(len(timing_element.findall(f".//{{{PREFIX_NAMESPACES['p']}}}seq/{{{PREFIX_NAMESPACES['p']}}}cTn/{{{PREFIX_NAMESPACES['p']}}}childTnLst")[0]))

def animation_directory(
        pptx_directory_path: Path,
        slide_number: int,
        uid: int,
        animation_name: str,
        animation_start: str = "click",
        animation_index: int | None = None,
        animation_delay: int | None = None,
        animation_attribs: dict | None = None
    ) -> None:
    """
    Sets the animation animation_name on the element with id uid of slide slide_number,
    at index animation_index relative to preexisting animations in the <p:timing> block
    of the extracted .pptx file directory pointed at by pptx_directory_path.
    The animation can optionally be fully or partially configured using the animation_attribs dict.
    The animation can optionally be configured to start on click or with/after the previous animation,
    delayed animation_delay milliseconds, and can be fully or partially configured using the animation_attribs dict.
    """
    slide_path = utils.get_slide_path(pptx_directory_path, slide_number)
    slide = ET.parse(slide_path).getroot()
    timing_element = get_timing_element(slide)
    extend_timing_element(timing_element, uid, animation_name, animation_start, animation_index, animation_delay, animation_attribs)
    # utils.save_xml(slide, slide_path)

def animation(
        pptx_path: Path,
        slide_number: int,
        uid: int,
        animation_name: str,
        animation_start: str = "click",
        animation_index: int | None = None,
        animation_delay: int | None = None,
        animation_attribs: dict | None = None
    ) -> None:
    """
    Sets the animation animation_name on the element with id uid of slide slide_number,
    at index animation_index relative to preexisting animations in the <p:timing> block
    of the .pptx file or extracted .pptx file directory pointed at by pptx_path.
    The animation can optionally be configured to start on click or with/after the previous animation,
    delayed animation_delay milliseconds, and can be fully or partially configured using the animation_attribs dict.
    """
    return utils.pptx_path_handler(pptx_path, animation_directory, [slide_number, uid, animation_name, animation_start, animation_index, animation_delay, animation_attribs])

def main():
    parser = argparse.ArgumentParser(description="Configures an animation on a specific slide.")
    parser.add_argument("-p", "--pptx-path", type=str, required=True, help="Path to a .pptx file or a directory corresponding to an extracted .pptx file.")
    parser.add_argument("-s", "--slide-number", type=int, required=True, help="Slides to modify, provided by its slide number (counting from 1).")
    parser.add_argument("-id", "--identifier", type=int, required=True, help="Unique identifier of the preexisting <p:spTree> child element to which the animation will be applied.")
    parser.add_argument("-an", "--animation-name", type=str, required=True, help=f"Name of the animation to be applied. Available animations are {', '.join(ANIMATIONS.keys())}")
    parser.add_argument("-as", "--animation-start", type=str, help="Start condition of the animation. Available options are \"click\" (on click), \"with\" (with previous) and \"after\" (after previous). If not provided, it is \"click\".")
    parser.add_argument("-ai", "--animation-index", type=int, help="Index in the <p:Timing> animation stack to insert the animation at. If not provided, the animation is appended at the top of the stack as the last animation of the slide. Out-of-range indexes get clamped so as to also insert the animation at the top of the stack. Index 0 corresponds to the first animation.")
    parser.add_argument("-ad", "--animation-delay", type=int, help="Animation delay in milliseconds. If not provided, there is no delay.")
    parser.add_argument("-at", "--animation-attribs", type=str, help="Attribute values of the animation to be applied, provided as a JSON string mapping each attribute to its value by attribute name. If an attribute is not provided, its default value will be used. Attributes are named exactly as according to the PresentationML specs.")
    args = parser.parse_args()
    arg_pptx_path = Path(args.pptx_path)
    arg_slide_number = args.slide_number
    arg_uid = args.identifier
    arg_animation_name = args.animation_name
    arg_animation_start = args.animation_start or "click"
    arg_animation_index = args.animation_index
    arg_animation_delay = args.animation_delay
    arg_animation_attribs = args.animation_attribs
    errors.error_validation_path_missing(arg_pptx_path)
    errors.error_validation_unavailable_animation(arg_animation_name, list(ANIMATIONS.keys()))
    errors.error_validation_unavailable_animation_start_condition(arg_animation_start)
    errors.error_validation_invalid_attribs_json(arg_animation_attribs)
    arg_animation_attribs = None if arg_animation_attribs is None else json.loads(arg_animation_attribs)
    errors.error_validation_extra_attribs(arg_animation_attribs, ANIMATIONS[arg_animation_name]["attribs"])
    errors.error_validation_mistyped_attribs(arg_animation_attribs, ANIMATIONS[arg_animation_name]["attribs"])
    errors.error_validation_invalid_attribs(arg_animation_attribs, ANIMATIONS[arg_animation_name]["attribs"])
    errors.error_validation_slide_numbers_out_of_range([arg_slide_number], utils.get_slide_count(arg_pptx_path))
    errors.error_validation_negative_numbers({"-ai (--animation_index)": arg_animation_index, "-ad (--animation_delay)": arg_animation_delay})
    errors.error_validation_unused_uid(arg_uid, utils.get_slide_uids(arg_pptx_path, arg_slide_number))
    animation(arg_pptx_path, arg_slide_number, arg_uid, arg_animation_name, arg_animation_index, arg_animation_delay, arg_animation_attribs)

if __name__ == "__main__":
    # main()
    animation(
        Path("/home/tomo/Documents/misc/pepetex/examples/tests/animations/anim"),
        5,
        0,
        "box_in"
    )