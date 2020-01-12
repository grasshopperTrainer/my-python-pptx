# encoding: utf-8

"""lxml custom element classes for DrawingML line-related XML elements."""

from __future__ import absolute_import, division, print_function, unicode_literals

from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrOneChoice, Choice,ZeroOrOne
from pptx.oxml.simpletypes import (
    ST_Coordinate,
    ST_CapType,
    ST_EdgeType,
    ST_EdgeAlignment,
    ST_LineEndType,
    ST_ReletiveSize,
)

class CT_PresetLineDashProperties(BaseOxmlElement):
    """`a:prstDash` custom element class"""

    val = OptionalAttribute("val", MSO_LINE_DASH_STYLE)

class _CT_TableCellEdge(BaseOxmlElement):
    """
    ''<a:ln{}>'' custom element class
    """
    w = RequiredAttribute('w', ST_Coordinate)
    cap = RequiredAttribute('cap', ST_CapType)
    cmpd = RequiredAttribute('cm    pd', ST_EdgeType)
    algn = RequiredAttribute('algn', ST_EdgeAlignment)

    _tag_seq = ('a:prstDash', 'a:headEnd', 'a:tailEnd')
    eg_fillProperties = ZeroOrOneChoice(
        (
            Choice("a:noFill"),
            Choice("a:solidFill"),
            Choice("a:gradFill"),
            Choice("a:blipFill"),
            Choice("a:pattFill"),
            Choice("a:grpFill"),
        ), successors=_tag_seq
    )
    eg_JointProperties = ZeroOrOneChoice(
        (
            Choice("a:bevel"),
            Choice("a:miter"),
            Choice("a:round"),
        ), successors=_tag_seq[1:]
    )

    prstDash = ZeroOrOne('a:prstDash', successors=('a:headEnd','a:bevel','a:miter','a:round','a:tailEnd'))
    headEnd = ZeroOrOne('a:headEnd', successors=_tag_seq[2:])
    tailEnd = ZeroOrOne('a:prstDash')
    del _tag_seq

class CT_TableCellEdgeLeft(_CT_TableCellEdge):
    pass
class CT_TableCellEdgeRight(_CT_TableCellEdge):
    pass
class CT_TableCellEdgeTop(_CT_TableCellEdge):
    pass
class CT_TableCellEdgeBottom(_CT_TableCellEdge):
    pass

class CT_LineHeadEnd(BaseOxmlElement):
    """
    ''<a:headEnd>''
    """
    type = RequiredAttribute('type', ST_LineEndType)
    w = OptionalAttribute('w',ST_ReletiveSize)
    len = OptionalAttribute('len',ST_ReletiveSize)

class CT_LineTailEnd(BaseOxmlElement):
    """
    ''<a:tailEnd>''
    """
    type = RequiredAttribute('type', ST_LineEndType)
    w = OptionalAttribute('w',ST_ReletiveSize)
    len = OptionalAttribute('len',ST_ReletiveSize)
