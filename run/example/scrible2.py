import pptx
import copy
from pptx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT
from pptx.parts.slide import SlidePart
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.packuri import PackURI
from pptx.util import Cm as cm
import pptx.virtual as vr


prsnt = pptx.Presentation()

slide = vr.Slide()
slide.shapes.clear_all()
slide.shapes.add_table(2,2,cm(5),cm(5),cm(1),cm(1))
slide.append_to(prsnt)
slide.append_to(prsnt)

prsnt.save('ccc.pptx')