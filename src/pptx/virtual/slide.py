from ..parts.slide import SlidePart
from ..oxml.slide import CT_Slide
from ..opc.packuri import PackURI
from ..opc.constants import RELATIONSHIP_TYPE as RT
from .. import Presentation
import copy

class Slide:
    _false_package = Presentation()

    def __init__(self, slide_layout_idx = 0):
        self._slide = self._false_package.slides.new_slide(self._false_package.slide_layouts[slide_layout_idx])

    def append_to(self, prsnt):
        """
        Adds this slide to a Presentation.
        :param prsnt:
        :return:
        """
        # find matching slide layout by name or set default as 0
        slide_layout = prsnt.slide_layouts[0]
        for sl in prsnt.slide_layouts:
            if self._slide.slide_layout.name == sl.name:
                slide_layout = sl

        slide = prsnt.slides.new_slide(slide_layout)
        slide._part._element = self._slide._part._element
        slide._element = self._slide._element

    def __getattr__(self, item):
        """
        Wires virtual slide to actual slide
        :param item:
        :return:
        """
        if hasattr(self._slide, item):
            return eval(f'self._slide.{item}')
