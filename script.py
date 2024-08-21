#region library
import clr 
import os
import sys
clr.AddReference("System")
import System

clr.AddReference("RevitServices")
import RevitServices
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import Autodesk
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from System.Collections.Generic import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.DB.Mechanical import *


from System.Windows import MessageBox
from System.IO import FileStream, FileMode, FileAccess
from System.Windows.Markup import XamlReader
#endregion

#region revit infor
# Get the directory path of the script.py & the Window.xaml
dir_path = os.path.dirname(os.path.realpath(__file__))
xaml_file_path = os.path.join(dir_path, "Window.xaml")

#Get UIDocument, Document, UIApplication, Application
uidoc = __revit__.ActiveUIDocument
uiapp = UIApplication(uidoc.Document.Application)
app = uiapp.Application
doc = uidoc.Document
activeView = doc.ActiveView
#endregion

#region method
class FilterDucts(ISelectionFilter):
    def AllowElement(self, element):
        if element.Category.Name == "Ducts": return True
        else: return False
             
    def AllowReference(self, reference, position):
        return True

class Utils:
    def create_duct_fitting(self, list_duct_ids):
        duct_0 = doc.GetElement(list_duct_ids[0])
        for i in range(1, len(list_duct_ids)):
            duct_i = doc.GetElement(list_duct_ids[i])
            self.create_union_fitting(duct_0, duct_i)
            duct_0 = duct_i
        
        return None

    def create_union_fitting (self, duct_1, duct_2):
        cM1 = duct_1.ConnectorManager
        cM2 = duct_2.ConnectorManager

        cS1 = cM1.Connectors
        cS2 = cM2.Connectors
        listCns = []

        for cn1 in cS1:
            for cn2 in cS2:
                p1 = cn1.Origin
                p2 = cn2.Origin
                distance = round(p1.DistanceTo(p2),0)
                if distance == 0:
                    listCns.append(cn1)
                    listCns.append(cn2)
                    break
        try:
            doc.Create.NewUnionFitting(listCns[0], listCns[1])
        except:
            pass
    
    def split_duct_from_start_point (self, duct, distance):
        
        try:
            ids = []
            location = duct.Location
            lc_line = location.Curve
            number = lc_line.Length / distance
            total = round(number,0) 
            i = 0
            while i < total :
                try:
                    p = self.find_point_from_start_point(lc_line, distance)
                    id = MechanicalUtils.BreakCurve(doc, duct.Id, p)
                    ids.append(id)
                    i+=1
                except:
                    break
        
            ids.append(duct.Id)
            self.create_duct_fitting(ids)
        except:
            pass
    
    def split_duct_from_end_point (self, duct, distance):
        try:
            location = duct.Location
            lc_line = location.Curve
            number = lc_line.Length / distance
            total = round(number,0)

            ids = [duct.Id]
            line = lc_line
            i = 0
            while i < total :
                try:
                    p = self.find_point_from_end_point(line, distance)
                    id = MechanicalUtils.BreakCurve(doc, duct.Id, p)
                    ids.append(id)

                    #reset data
                    duct = doc.GetElement(id)
                    new_location = duct.Location
                    line = new_location.Curve
                    i+=1
                except:
                    break
            self.create_duct_fitting(ids)
        except:
            pass

    def find_point_from_start_point(self, line, distance):
        sp = line.GetEndPoint(0)
        ep = line.GetEndPoint(1)
        dir = ep - sp
        tile = distance / dir.GetLength()

        x = tile * dir.X + sp.X
        y = tile * dir.Y + sp.Y
        z = tile * dir.Z + sp.Z

        return XYZ(x, y, z)

    def find_point_from_end_point(self, line, distance):
        sp = line.GetEndPoint(0)
        ep = line.GetEndPoint(1)
        dir =sp - ep
        tile = distance / dir.GetLength()

        x = tile * dir.X + ep.X
        y = tile * dir.Y + ep.Y
        z = tile * dir.Z + ep.Z

        return XYZ(x, y, z)

#endregion

#defind window
class WPFWindow:

    def load_window (self, list_duct):
        #import window from .xaml file path
        file_stream = FileStream(xaml_file_path, FileMode.Open, FileAccess.Read)
        window = XamlReader.Load(file_stream)

        #controls
        self.tb_distance = window.FindName("tb_Distance")
        self.cbb_rules = window.FindName("cbb_Rules")
        self.bt_Cancel = window.FindName("bt_Cancel")
        self.bt_OK = window.FindName("bt_Ok")
        
        #bindingdata
        self.list_duct = list_duct
        self.bindind_data()
        self.window = window
        return window


    def bindind_data (self):
        self.cbb_rules.ItemsSource =  ["From Start Point","From End Point"]
        self.bt_Cancel.Click += self.cancel_click
        self.bt_OK.Click += self.ok_click
        

    def ok_click(self, sender, e):
        rule = self.cbb_rules.SelectedItem
        text = self.tb_distance.Text
        distance = float(text)/304.8

        t = Transaction(doc, "slit duct")
        t.Start()

        for duct in self.list_duct:
            if str(rule).__contains__("Start"):
                Utils().split_duct_from_start_point(duct, distance)
            else:
                Utils().split_duct_from_end_point(duct, distance)

        t.Commit()
        self.window.Close()
        
    def cancel_click (self, sender, e):
        self.window.Close()
        
#select elements
class Main ():
    def get_list_duct (self):
        list_duct = []
        try:
            pick_objects = uidoc.Selection.PickObjects(ObjectType.Element, FilterDucts(), "Select Ducts")
            for r in pick_objects:
                duct = doc.GetElement(r)
                list_duct.append(duct)
        except:
            pass
        return list_duct
    
    def main_task(self):
        try:
            list_duct = self.get_list_duct()
            if len(list_duct) > 0 :
                window = WPFWindow().load_window(list_duct)
                window.ShowDialog()
        except Exception as e:
            MessageBox.Show(str(e), "Message")
        
        
if __name__ == "__main__":
    Main().main_task()
                
    
    






