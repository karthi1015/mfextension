# coding: utf8
from math import pi

from Autodesk.Revit.DB import Line, InsulationLiningBase
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit import Exceptions

from pyrevit import script
from pyrevit import forms
import rpw

import clr
import os
import os.path as op
import pickle as pl

import sys
import subprocess
import time

import struct

import rpw
from rpw import doc, uidoc, DB, UI

from System.Collections.Generic import List
from Autodesk.Revit.DB import *

import System

__doc__ = """Disconnect an object from it's connections.
"""
__title__ = "Disconnect Object"
__author__ = "Ed Green"

logger = script.get_logger()
uidoc = rpw.revit.uidoc
doc = rpw.revit.doc


from Autodesk.Revit.DB import Element, ConnectorManager, ConnectorSet, Connector, XYZ


def get_connector_manager(element):
    # type: (Element) -> ConnectorManager
    """Return element connector manager"""
    try:
        # Return ConnectorManager for pipes, ducts etc…
        return element.ConnectorManager
    except AttributeError:
        pass

    try:
        # Return ConnectorManager for family instances etc…
        return element.MEPModel.ConnectorManager
    except AttributeError:
        raise AttributeError("Cannot find connector manager in given element")


def get_connector_closest_to(connectors, xyz):
    # type: (ConnectorSet, XYZ) -> Connector
    """Get connector from connector set or any iterable closest to an XYZ point"""
    min_distance = float("inf")
    closest_connector = None
    for connector in connectors:
        distance = connector.Origin.DistanceTo(xyz)
        if distance < min_distance:
            min_distance = distance
            closest_connector = connector
    return closest_connector






class NoInsulation(ISelectionFilter):
    def AllowElement(self, elem):
        if isinstance(elem, InsulationLiningBase):
            return False
        try:
            get_connector_manager(elem)
            return True
        except AttributeError:
            return False

    def AllowReference(self, reference, position):
        return True


def disconnect_object():
    # Prompt user to select elements and points to connect
    try:
        with forms.WarningBar(title="Pick element to disconnet"):
            reference = uidoc.Selection.PickObject(ObjectType.Element, NoInsulation(), "Pick element to move")
    except Exceptions.OperationCanceledException:
        return False

    try:
		selected_element = doc.GetElement(reference)
		connectors = get_connector_manager(selected_element).Connectors
		for c in connectors:
			connectedTo = c.AllRefs
			
			for con in connectedTo:
				print str(con.Owner)
				
				c.DisconnectFrom(con)
				
				
				
			
			#print "Connected to:" str(connectedTo ) 
		return True
    except Exceptions.OperationCanceledException:
		print "Error"
		return True
	
	
	
	
	


t = Transaction(doc, 'Disconnect Object')
 
t.Start()

while disconnect_object():
    pass
	
t.Commit()	
