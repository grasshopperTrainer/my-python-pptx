from pptx.table import Table as _Table
from pptx.oxml.table import CT_Table
from pptx.oxml.shapes.graphfrm import CT_GraphicalObjectFrame
from pptx.oxml.slide import CT_Slide
from pptx.slide import SlideMaster
from pptx.opc.constants import RELATIONSHIP_TYPE as RT, CONTENT_TYPE as CT
from pptx.parts.slide import SlidePart
import pptx
import copy

class Table:
    def __init__(self,rows,cols, left, top, width, height):
        self.graphicFrame = CT_GraphicalObjectFrame.new_table_graphicFrame(-1,None,rows,cols,left,top,width, height)

    def append_to(self):
        pass


class Slide:
    def __init__(self, layout_idx=0):
        p = pptx.Presentation()
        l = p.slide_layouts[layout_idx]
        self.slide = p.slides.add_slide(l)

    def append_in(self, p):
        new_slide = copy.deepcopy(self.slide)
        # print(p.__dir__())
        # print(p.part.__dir__())
        new_slide.part._package = p.part.package
        # print(new_slide.part.related_parts)
        # print(new_slide.part._element.xml)
        # print(new_slide.slide_layout.__dir__())
        # print(new_slide.slide_layout.name)
        new_slide.shapes.add_table(2,2,0,0,pptx.util.Cm(10),pptx.util.Cm(10))
        # print(new_slide.shapes._spTree.xml)
        # exit()
        # print(new_slide.slide_layout.background)
        for i,sl in enumerate(p.slide_layouts):
            if sl.name == new_slide.slide_layout.name:
                new_slide.part.relate_to(sl.part,RT.SLIDE_LAYOUT)
                break
        # exit()
        k = p.part.relate_to(new_slide.part,reltype=RT.SLIDE)
        # print(k)
        # p.slides._sldIdLst.add_sldId(k)
        # print(p.slides[0].shapes._spTree.xml)
        # print(p.slides.part._next_slide_partname)
        # for i in p.part.rels.items():
        #     print(i)
        #     # print(i[1].__dir__())
        #     v = i[1]
        #     print(v.reltype, v._target, v.target_part, v.target_ref)
        # for i in p.part.parts.items():
        #     print(i)