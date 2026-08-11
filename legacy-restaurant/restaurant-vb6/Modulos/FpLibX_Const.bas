Attribute VB_Name = "FpLibX_Const"
' General error
Const ERROR_NONE = 0
Const ERROR_CREATION_FAILED = 1
Const ERROR_FUNCTION_FAILED = 2
Const ERROR_INVALID_PARAM = 3
Const ERROR_NOT_USED = 4
Const ERROR_DLLLOAD_FAILED = 5
Const ERROR_DLLLOAD_FAILED_DRV = 6
Const ERROR_DLLLOAD_FAILED_ALGO = 7
   
' Device error
Const ERROR_SYSLOAD_FAILED = 51     'system file load fail
Const ERROR_INITIALIZE_FAILED = 52  'chip initialize fail
Const ERROR_LINE_DROPPED = 53       'image data drop
Const ERROR_TIME_OUT = 54           'getliveimage timeout error
Const ERROR_DEVICE_NOT_FOUND = 55   'device not found
Const ERROR_DRVLOAD_FAILED = 56     'dll file load fail
Const ERROR_WRONG_IMAGE = 57         'wrong image
Const ERROR_LACK_OF_BANDWIDTH = 58   'USB Bandwith Lack Error
Const ERROR_DEV_ALREADY_OPEN = 59     'Device Exclusive access Error
Const ERROR_GETSN_FAILED = 60         'Fail to get Device Serial Number
Const ERROR_UNSUPPORTED_DEV = 61      ' Unsupported device
   
   
' Extract&verification error
Const ERROR_FEAT_NUMBER = 101               'too small number of minutiae
Const ERROR_INVALID_TEMPLATE_TYPE = 102     'wrong template type
Const ERROR_INVALID_TEMPLATE1 = 103         'error in decoding template 1
Const ERROR_INVALID_TEMPLATE2 = 104         'error in decoding template 2
Const ERROR_EXTRACT_FAIL = 105
Const ERROR_MATCH_FAIL = 106

' Template Format
Const TEMPLATE_FORMAT_ANSI378 = 256
Const TEMPLATE_FORMAT_SG400 = 512
Const TEMPLATE_FORMAT_ISO19794 = 768

