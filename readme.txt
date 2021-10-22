Instructions:
 For arguments:
  1. To pass an argument use format "-argumentName argumentValue"
  2. Only passing the "argumentValue" will save the argumentValue as UNKNOWARG
  3. Shuffling of arguments is possible, do not have to maintain sequence
  4. For running macro, module can be defined with "-module"(Ex. -module tags), but not mandatory, process will run in Design module as default
  5. Mandatory arguments: PDMSCommand, projectCode, mdb
Need Changes:
  1. Change path for 'automatedServiceFile' for RVM export process
  2. Change path for process monitor vbs
  3. Change path of log file(RVM)
  3. Change path of log file(MACRO)