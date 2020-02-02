# encoding: utf-8

"""
Objects shared by pptx modules.
"""

from __future__ import absolute_import, division, print_function, unicode_literals
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
import copy

class ElementProxy(object):
    """
    Base class for lxml element proxy classes. An element proxy class is one
    whose primary responsibilities are fulfilled by manipulating the
    attributes and child elements of an XML element. They are the most common
    type of class in python-pptx other than custom element (oxml) classes.
    """

    __slots__ = ("_element",)

    def __init__(self, element):
        self._element = element

    def __eq__(self, other):
        """
        Return |True| if this proxy object refers to the same oxml element as
        does *other*. ElementProxy objects are value objects and should
        maintain no mutable local state. Equality for proxy objects is
        defined as referring to the same XML element, whether or not they are
        the same proxy object instance.
        """
        if not isinstance(other, ElementProxy):
            return False
        return self._element is other._element

    def __ne__(self, other):
        if not isinstance(other, ElementProxy):
            return True
        return self._element is not other._element

    @property
    def element(self):
        """
        The lxml element proxied by this object.
        """
        return self._element


class ParentedElementProxy(ElementProxy):
    """
    Provides common services for document elements that occur below a part
    but may occasionally require an ancestor object to provide a service,
    such as add or drop a relationship. Provides the :attr:`_parent`
    attribute to subclasses and the public :attr:`parent` read-only property.
    """

    __slots__ = ("_parent",)

    def __init__(self, element, parent):
        super(ParentedElementProxy, self).__init__(element)
        self._parent = parent

    @property
    def parent(self):
        """
        The ancestor proxy object to this one. For example, the parent of
        a shape is generally the |SlideShapes| object that contains it.
        """
        return self._parent

    @property
    def part(self):
        """
        The package part containing this object
        """
        return self._parent.part


class PartElementProxy(ElementProxy):
    """
    Provides common members for proxy objects that wrap the root element of
    a part such as `p:sld`.
    """

    __slots__ = ("_part",)

    def __init__(self, element, part):
        super(PartElementProxy, self).__init__(element)
        self._part = part

    # def __deepcopy__(self, memodict={}):
    #     new_ins = self.__class__(self._element, copy.copy(self._part))
    #     return new_ins
    @property
    def part(self):
        """
        The package part containing this object
        """
        return self._part
