# encoding: utf-8

"""Graphic Frame shape and related objects.

A graphic frame is a common container for table, chart, smart art, and media
objects.
"""

from __future__ import absolute_import, division, print_function, unicode_literals

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.base import BaseShape
from pptx.table import Table
from pptx.util import Length

class GraphicFrame(BaseShape):
    """Container shape for table, chart, smart art, and media objects.

    Corresponds to a ``<p:graphicFrame>`` element in the shape tree.
    """

    @property
    def chart(self):
        """
        The |Chart| object containing the chart in this graphic frame. Raises
        |ValueError| if this graphic frame does not contain a chart.
        """
        if not self.has_chart:
            raise ValueError("shape does not contain a chart")
        return self.chart_part.chart

    @property
    def chart_part(self):
        """
        The |ChartPart| object containing the chart in this graphic frame.
        """
        rId = self._element.chart_rId
        chart_part = self.part.related_parts[rId]
        return chart_part

    @property
    def has_chart(self):
        """
        |True| if this graphic frame contains a chart object. |False|
        otherwise. When |True|, the chart object can be accessed using the
        ``.chart`` property.
        """
        return self._element.has_chart

    @property
    def has_table(self):
        """
        |True| if this graphic frame contains a table object. |False|
        otherwise. When |True|, the table object can be accessed using the
        ``.table`` property.
        """
        return self._element.has_table

    @property
    def shadow(self):
        """Unconditionally raises |NotImplementedError|.

        Access to the shadow effect for graphic-frame objects is
        content-specific (i.e. different for charts, tables, etc.) and has
        not yet been implemented.
        """
        raise NotImplementedError("shadow property on GraphicFrame not yet supported")

    @property
    def shape_type(self):
        """
        Unique integer identifying the type of this shape, e.g.
        ``MSO_SHAPE_TYPE.TABLE``.
        """
        if self.has_chart:
            return MSO_SHAPE_TYPE.CHART
        elif self.has_table:
            return MSO_SHAPE_TYPE.TABLE
        else:
            return None

    @property
    def table(self):
        """
        The |Table| object contained in this graphic frame. Raises
        |ValueError| if this graphic frame does not contain a table.
        """
        if not self.has_table:
            raise ValueError("shape does not contain a table")
        tbl = self._element.graphic.graphicData.tbl
        return Table(tbl, self)

    def orient(self, posx, posy, ref=0):
        """
        Reposition with reference vertice of bounding rectangle:
        0-------1
        /       /
        /   4   /
        /       /
        3-------4

        :param posx:
        :param posy:
        :param ref:
        :return:
        """
        if any([not isinstance(x, Length) for x in (posx, posy)]):
            raise TypeError


        if ref == 0:
            self.left, self.top = posx, posy
        elif ref == 1:
            self.left, self.top = posx - self.width, posy
        elif ref == 2:
            self.left, self.top = posx - self.width, posy - self.height
        elif ref == 3:
            self.left, self.top = posx, posy - self.height
        elif ref == 4:
            self.left, self.top = posx - self.width/2, posy - self.height/2
        else:
            raise TypeError