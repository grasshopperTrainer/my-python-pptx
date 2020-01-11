import pptx
import pptx.oxml.xmlchemy as xmlchemy

#
class CT_A(xmlchemy.BaseOxmlElement):
    b = xmlchemy.ZeroOrMore('a:bb')

    def add_b(self, v):
        return self._add_b(h=v)


# class CT_BooBoo(xmlchemy.BaseOxmlElement):
#     """
#     ``<a:tr>`` custom element class
#     """
#
#     tc = xmlchemy.ZeroOrMore("a:tc", successors=("a:extLst",))
#     h = xmlchemy.RequiredAttribute("h", int)
#
#     def add_tc(self):
#         """
#         Return a reference to a newly added minimal valid ``<a:tc>`` child
#         element.
#         """
#         return self._add_tc()
#
#     @property
#     def row_idx(self):
#         """Offset of this row in its table."""
#         return self.getparent().tr_lst.index(self)
    #
    # def _new_tc(self):
    #     return CT_TableCell.new()

a = CT_A()
# b = B()
print(a.add_b(10))

# from pptx.oxml.table import CT_Table, CT_TableRow
#
# a = CT_Table()
# b = a.add_tr(height=10)
# print(b)