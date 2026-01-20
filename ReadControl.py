# ReadControl.py - Utility script for reading control file.

# Suggested Enhancements:
#       Define class ReadControl.
#       Add method to verify that parameter exists and has the requisite number of variables.

import shlex

def help() :
    print("ReadControl.py - Utility script for reading control file.")
    print("\nUsage:")
    print("   import ReadControl   # then, call individual functions (see below)")
    print("\nFunctions:")
    print("\n    ReadControl.help()            # displays information about ReadControl.py   No return values.")
    print("\n    ReadControl.read(<path>, <return_lists>)")
    print("       path = full path to control file")
    print("       return_lists (optional) = boolean - defaults to False: return maximum one value per key, as string object.  If True, will return Disctionary values as List objects.")
    print("    Returns: <return_code> <dictionary>")
    print("       return_code: 0 = successful completion, any other value => error")
    print("       dictionary: dictionary object")
    print(" ")
    
#   function ReadControl.read(<path>, <return_lists>) - Read Control File and load values into a dictionary object.
def read(control_path, return_lists=False) :
    try:
        input_file = open(control_path, 'r')
    except FileNotFoundError:
        print(f"Control File \"{control_path}\" not found.")
        return 1, None
        
    parms_dict = {}
    for line in input_file:
        parm_tokens = shlex.split(line.strip())     # Development note: need to catch ValueError: No closing quotation - in case the Control File has unbalanced quotation marks. 
        if return_lists :
            if len(parm_tokens) > 1:
               return_list = parm_tokens[1:]
               parms_dict[parm_tokens[0]] = return_list
            elif len(parm_tokens) == 1:
               return_list = []
               parms_dict[parm_tokens[0]] = return_list
        else :
            if len(parm_tokens) > 1:
                parms_dict[parm_tokens[0]] = parm_tokens[1]
            elif len(parm_tokens) == 1:
                parms_dict[parm_tokens[0]] = None
            
    if len(parms_dict) < 1 :
        print(f"Did not find any parameters in Control File \"{control_path}\".")
        return 1, None
            
    input_file.close()
    return 0, parms_dict
    