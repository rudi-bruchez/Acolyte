import getpass, os, os.path, uno

class Session():
    @staticmethod
    def substitute(var_name):
        ctx = uno.getComponentContext()
        ps = ctx.getServiceManager().createInstanceWithContext(
            'com.sun.star.util.PathSubstitution', ctx)
        return ps.getSubstituteVariableValue(var_name)
    
    @staticmethod
    def Share():
        inst = uno.fileUrlToSystemPath(Session.substitute("$(prog)"))
        return os.path.normpath(inst.replace('program', "Share"))
    
    @staticmethod
    def SharedScripts():
        return ''.join([Session.Share(), os.sep, "Scripts"])
    
    @staticmethod
    def SharedPythonScripts():
        return ''.join([Session.SharedScripts(), os.sep, 'python'])
    
    @property  # alternative to '$(username)' variable
    def UserName(self): return getpass.getuser()
    
    @property
    def UserProfile(self):
        return uno.fileUrlToSystemPath(Session.substitute("$(user)"))
    
    @property
    def UserScripts(self):
        return ''.join([self.UserProfile, os.sep, 'Scripts'])
    
    @property
    def UserPythonScripts(self):
        return ''.join([self.UserScripts, os.sep, "python"])