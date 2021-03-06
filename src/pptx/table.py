# encoding: utf-8

"""Table-related objects such as Table and Cell."""

from __future__ import absolute_import, division, print_function, unicode_literals

from pptx.compat import is_integer
from pptx.dml.fill import FillFormat
from pptx.oxml.table import TcRange,CT_TableCell
from pptx.oxml.dml.line import CT_TableCellEdgeLeft, CT_TableCellEdgeRight, CT_TableCellEdgeTop, CT_TableCellEdgeBottom
from pptx.oxml.xmlchemy import serialize_for_reading
from pptx.shapes import Subshape
from pptx.text.text import TextFrame
from pptx.util import lazyproperty
from pptx.util import Length,Emu
from pptx.dml.line import LineFormat
import copy

class Table(object):
    """A DrawingML table object.

    Not intended to be constructed directly, use
    :meth:`.Slide.shapes.add_table` to add a table to a slide.
    """

    def __init__(self, tbl, graphic_frame):
        super(Table, self).__init__()
        self._tbl = tbl
        self._graphic_frame = graphic_frame

    def cell(self, row_idx, col_idx):
        """Return cell at *row_idx*, *col_idx*.

        Return value is an instance of |_Cell|. *row_idx* and *col_idx* are
        zero-based, e.g. cell(0, 0) is the top, left cell in the table.
        """
        return _Cell(self._tbl.tc(row_idx, col_idx), self)

    @lazyproperty
    def columns(self):
        """
        Read-only reference to collection of |_Column| objects representing
        the table's columns. |_Column| objects are accessed using list
        notation, e.g. ``col = tbl.columns[0]``.
        """
        return _ColumnCollection(self._tbl, self)

    @property
    def first_col(self):
        """
        Read/write boolean property which, when true, indicates the first
        column should be formatted differently, as for a side-heading column
        at the far left of the table.
        """
        return self._tbl.firstCol

    @first_col.setter
    def first_col(self, value):
        self._tbl.firstCol = value

    @property
    def first_row(self):
        """
        Read/write boolean property which, when true, indicates the first
        row should be formatted differently, e.g. for column headings.
        """
        return self._tbl.firstRow

    @first_row.setter
    def first_row(self, value):
        self._tbl.firstRow = value

    @property
    def horz_banding(self):
        """
        Read/write boolean property which, when true, indicates the rows of
        the table should appear with alternating shading.
        """
        return self._tbl.bandRow

    @horz_banding.setter
    def horz_banding(self, value):
        self._tbl.bandRow = value

    def iter_cells(self):
        """Generate _Cell object for each cell in this table.

        Each grid cell is generated in left-to-right, top-to-bottom order.
        """
        return (_Cell(tc, self) for tc in self._tbl.iter_tcs())
    def iter_real_cells(self):
        """
        Return cells with no merged
        :return:
        """
        return (_Cell(tc, self) for tc in self._tbl.iter_tcs() if not tc.is_spanned)

    @property
    def last_col(self):
        """
        Read/write boolean property which, when true, indicates the last
        column should be formatted differently, as for a row totals column at
        the far right of the table.
        """
        return self._tbl.lastCol

    @last_col.setter
    def last_col(self, value):
        self._tbl.lastCol = value

    @property
    def last_row(self):
        """
        Read/write boolean property which, when true, indicates the last
        row should be formatted differently, as for a totals row at the
        bottom of the table.
        """
        return self._tbl.lastRow

    @last_row.setter
    def last_row(self, value):
        self._tbl.lastRow = value

    def notify_height_changed(self):
        """
        Called by a row when its height changes, triggering the graphic frame
        to recalculate its total height (as the sum of the row heights).
        """
        new_table_height = sum([row.height for row in self.rows])
        self._graphic_frame.height = new_table_height

    def notify_width_changed(self):
        """
        Called by a column when its width changes, triggering the graphic
        frame to recalculate its total width (as the sum of the column
        widths).
        """
        new_table_width = sum([col.width for col in self.columns])
        self._graphic_frame.width = new_table_width

    @property
    def part(self):
        """
        The package part containing this table.
        """
        return self._graphic_frame.part

    @lazyproperty
    def rows(self):
        """
        Read-only reference to collection of |_Row| objects representing the
        table's rows. |_Row| objects are accessed using list notation, e.g.
        ``col = tbl.rows[0]``.
        """
        return _RowCollection(self._tbl, self)

    @property
    def vert_banding(self):
        """
        Read/write boolean property which, when true, indicates the columns
        of the table should appear with alternating shading.
        """
        return self._tbl.bandCol

    @vert_banding.setter
    def vert_banding(self, value):
        self._tbl.bandCol = value

    def add_row(self, n=1, height=None):
        """
        Adds number of rows into the table.
        :param n: number of rows to add
        :param height: if not given height of last row is used
        :return:
        """
        if not isinstance(n, int):
            raise
        if height is None:
            height = self.rows[-1].height
        elif not isinstance(height, Length):
            raise TypeError('Must feed one of pptx.util.Length types')

        trs = []
        for i in range(n):
            new_tr = self._tbl.add_tr(height)
            for col in self._tbl.tblGrid.gridCol_lst:
                trs.append(new_tr._add_tc())

        if len(trs) == 1:
            return trs[0]
        else:
            return trs

    def add_column(self, n=1, width=None):
        """
        Adds number of columns into the table
        :param n: number of columns to add
        :param width: if not given width of last column is used
        :return:
        """
        if not isinstance(n, int):
            raise
        if width is None:
            width = self.columns[-1].width
        elif not isinstance(width, Length):
            raise TypeError('Must feed one of pptx.util.Length types')

        for i in range(n):
            self._tbl.tblGrid.add_gridCol(width)
            for row in self._tbl.tr_lst:
                row.add_tc()

    def delete_row(self, idx=-1):
        """
        Deletes row of given index out from the table.
        :param idx: index of row to delete
        :return:
        """
        if not isinstance(idx, int):
            raise TypeError('Must feet <int>')
        self._tbl.tr_lst[idx].delete()

    def delete_column(self, idx=-1):
        """
        Deletes column of given index out from the table.
        :param idx: index of column to delete
        :return:
        """
        if not isinstance(idx, int):
            raise TypeError('Must feet <int>')
        for row in self._tbl.tr_lst:
            row.tc_lst[idx].delete()
        self._tbl.tblGrid.gridCol_lst[idx].delete()

    def join_table(self, to_join, pos=0, trim=True, remove_joined=True):
        """
        Joins another table at the given position.
        *pos* can be given as integer or string:
        {'left':0,'right':1,'bottom':2,'top':3}

        :param to_join: table to be joined
        :param pos: tells where to append table
        :param remove_joined: default is set to remove after join
        :return:
        """
        # TODO need position anchoring?

        if not isinstance(to_join, Table):
            raise TypeError('param *table* should be <Table> type')

        if isinstance(pos, int):
            if not (pos < 4 and pos >= 0):
                raise ValueError
        elif isinstance(pos, str):
            str_int_dict = {'left':0,'right':1,'bottom':2,'top':3}
            try:
                pos = str_int_dict[pos]
            except:
                raise ValueError(f'pos should be one of (left, right, bottom, top)')
        else:
            raise TypeError('param *pos* should be either <str> or <int>')

        that_table = copy.deepcopy(to_join)
        r_diff = len(self.rows) - len(to_join.rows)
        c_diff = len(self.columns) - len(to_join.columns)

        if pos in (0,1):
            if r_diff == 0:
                pass
            elif r_diff > 0:
                if trim:
                    for i in range(r_diff):
                        self.delete_row(-1)
                else:
                    that_table.add_row(r_diff)

            elif r_diff < 0:
                if trim:
                    for i in range(r_diff):
                        that_table.delete_row(-1)
                else:
                    self.add_row(r_diff)

            if pos == 0:
                for gc in that_table._tbl.tblGrid.gridCol_lst:
                    self._tbl.tblGrid.append(gc)

                for this_tr, that_tr in zip(self._tbl.tr_lst, that_table._tbl.tr_lst):
                    for tc in that_tr.tc_lst:
                        this_tr.append(tc)

            elif pos == 1:
                for gc in reversed(that_table._tbl.tblGrid.gridCol_lst):
                    self._tbl.tblGrid.insert(0, gc)

                for this_tr, that_tr in zip(self._tbl.tr_lst, that_table._tbl.tr_lst):
                    for tc in reversed(that_tr.tc_lst):
                        this_tr.insert(0, tc)

        if pos in (2,3):
            # graft table
            if c_diff == 0:
                pass
            elif c_diff > 0:
                if trim:
                    for i in range(c_diff):
                        self.delete_column(-1)
                else:
                    that_table.add_column(c_diff)
            elif c_diff < 0:
                if trim:
                    for i in range(-c_diff):
                        that_table.delete_column(-1)
                else:
                    self.add_column(-c_diff)

            if pos == 2:
                for tr in that_table._tbl.tr_lst:
                    self._tbl.append(tr)
            elif pos == 3:
                for tr in reversed(that_table._tbl.tr_lst):
                    self._tbl.insert(0,tr)


        if remove_joined:
            ss = to_join._graphic_frame._parent
            ss._spTree.remove(to_join._graphic_frame._element)

    def cell_idx(self, cell):
        if not isinstance(cell, _Cell):
            raise

        tc = cell._tc
        tr = tc.getparent()

        for i, c in enumerate(tr.tc_lst):
            if c == tc:
                col_idx = i
                break

        row_idx = None
        for i, r in enumerate(self._tbl.tr_lst):
            if r == tr:
                row_idx = i
                break

        return row_idx, col_idx


class _CellEdge:
    def __init__(self, tc):
        self._tc = tc

    @property
    def left(self):
        if self._tc.tcPr.lnL is None:
            self._add_default_edge('L')
        return _Edge(self._tc.tcPr, self._tc.get_or_add_lnL())
    @property
    def right(self):
        if self._tc.tcPr.lnR is None:
            self._add_default_edge('R')
        return _Edge(self._tc.tcPr, self._tc.get_or_add_lnR())
    @property
    def top(self):
        if self._tc.tcPr.lnT is None:
            self._add_default_edge('T')
        return _Edge(self._tc.tcPr, self._tc.get_or_add_lnT())
    @property
    def bottom(self):
        if self._tc.tcPr.lnB is None:
            self._add_default_edge('B')
        return _Edge(self._tc.tcPr, self._tc.tcPr.get_or_add_lnB())

    def _add_default_edge(self, pos):
        ln = eval(f"self._tc.tcPr._add_ln{pos}(w=12700, cap='flat', cmpd='sng', algn='ctr')")
        sf = ln._add_solidFill()
        srgb = sf._add_srgbClr(val='FFFFFF')
        srgb._add_alpha(val=1.0)
        ln._add_round()
        ln._add_prstDash(val=1)
        ln._add_headEnd(type='none', w='med', len='med')
        ln._add_tailEnd(type='none', w='med', len='med')
        return ln

    def solid(self, r,g,b,a):
        pass

class _Edge:
    def __init__(self, tcPr, lnx):
        self._tcPr = tcPr
        self._lnx = lnx

    def solid(self, r,g,b,a=1.0):
        pass



# class _CellEdgeLeft(_Edge):
#     pass
# class _CellEdgeRight(_Edge):
#     pass
# class _CellEdgeTop(_Edge):
#     pass
# class _CellEdgeBottom(_Edge):
#     pass


class _Cell(Subshape):
    """Table cell"""

    def __init__(self, tc, parent):
        super(_Cell, self).__init__(parent)
        self._tc = tc

    def __eq__(self, other):
        """|True| if this object proxies the same element as *other*.

        Equality for proxy objects is defined as referring to the same XML
        element, whether or not they are the same proxy object instance.
        """
        if not isinstance(other, type(self)):
            return False
        return self._tc is other._tc

    def __ne__(self, other):
        if not isinstance(other, type(self)):
            return True
        return self._tc is not other._tc

    @lazyproperty
    def fill(self):
        """
        |FillFormat| instance for this cell, providing access to fill
        properties such as foreground color.
        """
        tcPr = self._tc.get_or_add_tcPr()
        return FillFormat.from_fill_parent(tcPr)

    @lazyproperty
    def edge(self):
        return _CellEdge(self._tc)

    @property
    def is_merge_origin(self):
        """True if this cell is the top-left grid cell in a merged cell."""
        return self._tc.is_merge_origin

    @property
    def is_spanned(self):
        """True if this cell is spanned by a merge-origin cell.

        A merge-origin cell "spans" the other grid cells in its merge range,
        consuming their area and "shadowing" the spanned grid cells.

        Note this value is |False| for a merge-origin cell. A merge-origin
        cell spans other grid cells, but is not itself a spanned cell.
        """
        return self._tc.is_spanned
    @property
    def margin(self):
        """
        Return all margin ordered left, right, top, bottom
        :return:
        """
        return self.margin_left, self.margin_right, self.margin_top, self.margin_bottom

    @margin.setter
    def margin(self, margin):
        """
        Sets all margins with equal value
        :param margin: margin for all four side
        :return:
        """
        self.margin_left = margin
        self.margin_bottom = margin
        self.margin_top = margin
        self.margin_bottom = margin

    @property
    def margin_left(self):
        """
        Read/write integer value of left margin of cell as a |Length| value
        object. If assigned |None|, the default value is used, 0.1 inches for
        left and right margins and 0.05 inches for top and bottom.
        """
        return self._tc.marL

    @margin_left.setter
    def margin_left(self, margin_left):
        self._validate_margin_value(margin_left)
        self._tc.marL = margin_left

    @property
    def margin_right(self):
        """
        Right margin of cell.
        """
        return self._tc.marR

    @margin_right.setter
    def margin_right(self, margin_right):
        self._validate_margin_value(margin_right)
        self._tc.marR = margin_right

    @property
    def margin_top(self):
        """
        Top margin of cell.
        """
        return self._tc.marT

    @margin_top.setter
    def margin_top(self, margin_top):
        self._validate_margin_value(margin_top)
        self._tc.marT = margin_top

    @property
    def margin_bottom(self):
        """
        Bottom margin of cell.
        """
        return self._tc.marB

    @margin_bottom.setter
    def margin_bottom(self, margin_bottom):
        self._validate_margin_value(margin_bottom)
        self._tc.marB = margin_bottom

    def merge(self, other_cell):
        """Create merged cell from this cell to *other_cell*.

        This cell and *other_cell* specify opposite corners of the merged
        cell range. Either diagonal of the cell region may be specified in
        either order, e.g. self=bottom-right, other_cell=top-left, etc.

        Raises |ValueError| if the specified range already contains merged
        cells anywhere within its extents or if *other_cell* is not in the
        same table as *self*.
        """
        tc_range = TcRange(self._tc, other_cell._tc)

        if not tc_range.in_same_table:
            raise ValueError("other_cell from different table")
        if tc_range.contains_merged_cell:
            raise ValueError("range contains one or more merged cells")

        tc_range.move_content_to_origin()

        row_count, col_count = tc_range.dimensions

        for tc in tc_range.iter_top_row_tcs():
            tc.rowSpan = row_count
        for tc in tc_range.iter_left_col_tcs():
            tc.gridSpan = col_count
        for tc in tc_range.iter_except_left_col_tcs():
            tc.hMerge = True
        for tc in tc_range.iter_except_top_row_tcs():
            tc.vMerge = True

    @property
    def span_height(self):
        """int count of rows spanned by this cell.

        The value of this property may be misleading (often 1) on cells where
        `.is_merge_origin` is not |True|, since only a merge-origin cell
        contains complete span information. This property is only intended
        for use on cells known to be a merge origin by testing
        `.is_merge_origin`.
        """
        return self._tc.rowSpan

    @property
    def span_width(self):
        """int count of columns spanned by this cell.

        The value of this property may be misleading (often 1) on cells where
        `.is_merge_origin` is not |True|, since only a merge-origin cell
        contains complete span information. This property is only intended
        for use on cells known to be a merge origin by testing
        `.is_merge_origin`.
        """
        return self._tc.gridSpan

    def split(self):
        """Remove merge from this (merge-origin) cell.

        The merged cell represented by this object will be "unmerged",
        yielding a separate unmerged cell for each grid cell previously
        spanned by this merge.

        Raises |ValueError| when this cell is not a merge-origin cell. Test
        with `.is_merge_origin` before calling.
        """
        if not self.is_merge_origin:
            raise ValueError(
                "not a merge-origin cell; only a merge-origin cell can be sp" "lit"
            )

        tc_range = TcRange.from_merge_origin(self._tc)

        for tc in tc_range.iter_tcs():
            tc.rowSpan = tc.gridSpan = 1
            tc.hMerge = tc.vMerge = False

    @property
    def text(self):
        """Unicode (str in Python 3) representation of cell contents.

        The returned string will contain a newline character (``"\\n"``) separating each
        paragraph and a vertical-tab (``"\\v"``) character for each line break (soft
        carriage return) in the cell's text.

        Assignment to *text* replaces all text currently contained in the cell. A
        newline character (``"\\n"``) in the assigned text causes a new paragraph to be
        started. A vertical-tab (``"\\v"``) character in the assigned text causes
        a line-break (soft carriage-return) to be inserted. (The vertical-tab character
        appears in clipboard text copied from PowerPoint as its encoding of
        line-breaks.)

        Either bytes (Python 2 str) or unicode (Python 3 str) can be assigned. Bytes can
        be 7-bit ASCII or UTF-8 encoded 8-bit bytes. Bytes values are converted to
        unicode assuming UTF-8 encoding (which correctly decodes ASCII).
        """
        return self.text_frame.text

    @text.setter
    def text(self, text):
        self.text_frame.text = text

    @property
    def text_frame(self):
        """
        |TextFrame| instance containing the text that appears in the cell.
        """
        txBody = self._tc.get_or_add_txBody()
        return TextFrame(txBody, self)

    @property
    def vertical_anchor(self):
        """Vertical alignment of this cell.

        This value is a member of the :ref:`MsoVerticalAnchor` enumeration or
        |None|. A value of |None| indicates the cell has no explicitly
        applied vertical anchor setting and its effective value is inherited
        from its style-hierarchy ancestors.

        Assigning |None| to this property causes any explicitly applied
        vertical anchor setting to be cleared and inheritance of its
        effective value to be restored.
        """
        return self._tc.anchor

    @vertical_anchor.setter
    def vertical_anchor(self, mso_anchor_idx):
        self._tc.anchor = mso_anchor_idx

    @staticmethod
    def _validate_margin_value(margin_value):
        """
        Raise ValueError if *margin_value* is not a positive integer value or
        |None|.
        """
        if not is_integer(margin_value) and margin_value is not None:
            tmpl = "margin value must be integer or None, got '%s'"
            raise TypeError(tmpl % margin_value)

    def coordinate(self, idx):
        """
        Return coordinate of cell area ordered starting from top_left going clockwise.
        :param idx: 0,1,2,3 and 4 for CENTER
        :return:
        """
        if not isinstance(idx, int):
            raise TypeError
        row_idx, col_idx = self._parent.cell_idx(self)
        tbl = self._tc.getparent().getparent()
        # table is inside graphic frame. Graphic frame is a thing that's inside slide.
        gf = tbl.getparent().getparent().getparent()
        top_left = [gf.x,gf.y]
        for tr in tbl.tr_lst[:row_idx]:
            top_left[1] += tr.h
        for gc in tbl.tblGrid.gridCol_lst[:col_idx]:
            top_left[0] += gc.w

        if idx == 0:
            return Emu(top_left[0]), Emu(top_left[1])
        elif idx == 1:
            return Emu(top_left[0] + self.width), Emu(top_left[1])
        elif idx == 2:
            return Emu(top_left[0]) + self.width, Emu(top_left[1] + self.height)
        elif idx == 3:
            return Emu(top_left[0]), Emu(top_left[1] + self.height)
        elif idx == 4:
            return Emu(top_left[0] + self.width/2), Emu(top_left[1] + self.height/2)
        else:
            raise ValueError

    @property
    def width(self):
        tbl = self._tc.getparent().getparent()
        col_idx = None
        for i,tc in enumerate(self._tc.getparent().tc_lst):
            if self._tc == tc:
                col_idx = i
                break
        if self._tc.is_merge_origin:
            gc = self._tc.getparent().getparent().tblGrid.gridCol_lst[col_idx]
            w = 0
            for i in range(self._tc.gridSpan):
                w += gc.w
                gc = gc.getnext()
            return w
        else:
            return tbl.tblGrid.gridCol_lst[col_idx].w

    @property
    def height(self):
        tr = self._tc.getparent()
        if self._tc.is_merge_origin:
            h = 0
            for i in range(self._tc.rowSpan):
                h += tr.h
                tr = tr.getnext()
            return h
        else:
            return tr.h

class _Column(Subshape):
    """Table column"""

    def __init__(self, gridCol, parent):
        super(_Column, self).__init__(parent)
        self._gridCol = gridCol

    @property
    def width(self):
        """
        Width of column in EMU.
        """
        return self._gridCol.w

    @width.setter
    def width(self, width):
        self._gridCol.w = width
        self._parent.notify_width_changed()

    @property
    def cells(self):
        col_idx = self._gridCol.getparent().index(self._gridCol)
        tbl = self._gridCol.getparent().getparent()
        tcs = []
        for r in tbl.tr_lst:
            tcs.append(r.tc_lst[col_idx])
        return _CellCollection(tcs, self)

class _Row(Subshape):
    """Table row"""

    def __init__(self, tr, parent):
        super(_Row, self).__init__(parent)
        self._tr = tr

    @property
    def cells(self):
        """
        Read-only reference to collection of cells in row. An individual cell
        is referenced using list notation, e.g. ``cell = row.cells[0]``.
        """
        return _CellCollection(self._tr.tc_lst, self)

    @property
    def height(self):
        """
        Height of row in EMU.
        """
        return self._tr.h

    @height.setter
    def height(self, height):
        self._tr.h = height
        self._parent.notify_height_changed()


class _CellCollection(Subshape):
    """Horizontal sequence of row cells"""

    def __init__(self, tcs, parent):
        """
        Wrapper class for indexing cells from row or column.
        :param tr_tc: <CT_TableRow> or <CT_TableCol>
        :param parent: <_Row> or <_Column>
        """
        super(_CellCollection, self).__init__(parent)
        self._tcs = tcs

    def __getitem__(self, idx):
        """Provides indexed access, (e.g. 'cells[0]')."""
        if isinstance(idx, int):
            return _Cell(self._tcs[idx], self)

        elif isinstance(idx, slice):
            start = 0 if idx.start is None else idx.start
            stop = len(self._tcs) if idx.stop is None else idx.stop
            step = 1 if idx.step is None else idx.step
            cc = _CellCollection([self._tcs[i] for i in range(start, stop, step)], self)
            return cc
        else:
            raise TypeError

    def __iter__(self):
        """Provides iterability."""
        return (_Cell(tc, self) for tc in self._tcs)

    def __len__(self):
        """Supports len() function (e.g. 'len(cells) == 1')."""
        return len(self._tcs)


class _ColumnCollection(Subshape):
    """Sequence of table columns."""

    def __init__(self, tbl, parent):
        super(_ColumnCollection, self).__init__(parent)
        self._tbl = tbl

    def __getitem__(self, idx):
        """
        Provides indexed access, (e.g. 'columns[0]').
        """
        if idx < 0:
            idx = len(self._tbl.tblGrid.gridCol_lst)-1+idx

        if idx < 0 or idx >= len(self._tbl.tblGrid.gridCol_lst):
            msg = "column index [%d] out of range" % idx
            raise IndexError(msg)
        return _Column(self._tbl.tblGrid.gridCol_lst[idx], self)

    def __len__(self):
        """
        Supports len() function (e.g. 'len(columns) == 1').
        """
        return len(self._tbl.tblGrid.gridCol_lst)

    def notify_width_changed(self):
        """
        Called by a column when its width changes. Pass along to parent.
        """
        self._parent.notify_width_changed()


class _RowCollection(Subshape):
    """Sequence of table rows"""

    def __init__(self, tbl, parent):
        super(_RowCollection, self).__init__(parent)
        self._tbl = tbl

    def __getitem__(self, idx):
        """
        Provides indexed access, (e.g. 'rows[0]').
        """
        if idx < 0:
            idx = len(self)-1+idx

        if idx < 0 or idx >= len(self):
            msg = "row index [%d] out of range" % idx
            raise IndexError(msg)
        return _Row(self._tbl.tr_lst[idx], self)

    def __len__(self):
        """
        Supports len() function (e.g. 'len(rows) == 1').
        """
        return len(self._tbl.tr_lst)

    def notify_height_changed(self):
        """
        Called by a row when its height changes. Pass along to parent.
        """
        self._parent.notify_height_changed()
