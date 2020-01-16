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

        print(p.part._package.__dir__())
        package = p.part._package
        print(package._rels)
        main_doc_part = package.main_document_part
        media_part = package._media_parts
        image_part = package._image_parts
        print(main_doc_part.__dir__())
        print(main_doc_part._partname)
        print(main_doc_part)
        print(main_doc_part._presentation)
        print(new_slide.part.partname)
        print(package.parts)

        # print(media_part.__dir__())
        # print(image_part.__dir__())
        #
        # for i in media_part:
        #     print(i)
        # for i in image_part:
        #     print(i)

        # for i in p.part._package.parts:
        #     print(i)
        # print()
        # p.slides.add_slide(p.slide_layouts[0])
        #
        # for i in p.part._package.parts:
        #     print(i)

        # print(p.part.add_slide())



        # p.slides.append(new_slide)

