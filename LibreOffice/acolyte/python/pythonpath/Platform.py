# -*- coding: utf-8 -*-
import os, platform
class Platform():
    @property
    def ComputerName(self): return platform.node()
    @property
    def DirSeparator(self): return os.sep
    @property
    def isLinux(self): return (self.OSName=='Linux')
    @property
    def isMacOSX(self): return (self.OSName=='Darwin')
    @property
    def isWindows(self): return (self.OSName=='Windows')
    @property
    def OSName(self): return platform.system()
    @property
    def PathDelimiter(self): return os.pathsep