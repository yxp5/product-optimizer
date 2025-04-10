# META classes
# Author: yxp5

import encode

class Property:
    
    def __init__(self, name):
        self.name = name
        self.uid = self.UID()
    
    def UID(self):
        return encode.sha512(f"{self.name}")[:20]
    
    def __str__(self):
        return self.name

class Feature:
    
    def __init__(self, name, prop, desc):
        self.name = name
        self.prop = prop if prop else Property("None")
        self.desc = desc
        self.uid = self.UID()
    
    def Info(self, tab=False):
        description = self.desc if self.desc else "None"
        
        if tab:
            info = f"\tFeature ({self.name}) has property: [{self.prop}]\n\tDescription: {description}"
        else:
            info = f"Feature ({self.name}) has property: [{self.prop}]\nDescription: {description}"
        return info
    
    def UID(self):
        return encode.sha512(self.Info())[:20]

    def __str__(self):
        return self.Info()

class Product:
    
    def __init__(self, name, req):
        self.name = name
        self.req = req
        self.uid = self.UID()
    
    def UpdReq(self, feature, need):
        self.req.update({feature: need})
    
    def DelReq(self, feature):
        self.req.pop(feature)

    def Info(self):
        info = f"[{self.name}]\nRequirements:\n"
        
        for feat, need in self.req.items():
            info += f"Feature: {feat}, Need: {need}\n"
        
        return info
    
    def UID(self):
        return encode.sha512(f"{self.name}")[:20]
    
    def __str__(self):
        return self.name






