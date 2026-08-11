Attribute VB_Name = "ModuloHardKey"
'Declare Function HARDkey Lib "hkey-w32.dll" (ByVal buf As String) As Long
'
'Public sBox1 As Variant
'Public sBox2 As Variant
'Public password As String
'
'Public Sub InitSB()
'
'password = "bBkkEvEVToeSgkNQ"
'
'sBox1 = Array(&HF8, &HA0, &H32, &HEB, &HC0, &HF2, &HD8, &H16, &H70, &HE, &H22, &HDA, &H44, &H85, &HC2, &H8D, _
'&H68, &H74, &H7E, &H3F, &H54, &H31, &HE2, &H5B, &H38, &H7C, &HAA, &HC6, &H11, &H3C, &H34, &H83, _
'&H58, &H73, &H1E, &HB8, &H6, &HE1, &H1C, &HA4, &H60, &HF6, &H8E, &HDC, &H4, &HC1, &H30, &HF3, _
'&HC8, &HF9, &HFC, &HAF, &H88, &HB1, &H9E, &HB5, &HE8, &HE4, &HB, &H4B, &H26, &HD9, &HD2, &HB4, _
'&H23, &H91, &H24, &H25, &HB0, &HA7, &HEC, &HA6, &HA, &H2E, &HD0, &H97, &H1, &HED, &H9C, &H49, _
'&H4F, &H12, &H80, &HCA, &H2, &HE3, &H19, &H5D, &HE0, &HE6, &H4C, &H3B, &H50, &H27, &H76, &HAC, _
'&H0, &H6E, &HF0, &H6F, &H10, &H4E, &HBA, &H3D, &H5C, &HBB, &H57, &HDF, &H8C, &H98, &H40, &HBF, _
'&HA8, &H3, &H9, &HFF, &H48, &H5F, &H7, &H7A, &H61, &HEA, &H94, &H2D, &HC, &HA1, &H62, &H4A, _
'&H28, &HBC, &H79, &H8A, &HC4, &H9A, &H5, &H87, &H1D, &HE7, &HF4, &H75, &H6D, &H41, &HD6, &H29, _
'&H6C, &H3A, &HD4, &H35, &H2C, &H90, &H6A, &H2F, &H64, &HCF, &H8F, &H45, &H96, &HC7, &HA2, &H77, _
'&H42, &HF5, &HCC, &HD5, &H2B, &H71, &H56, &HC9, &H51, &HEF, &H20, &HE5, &H52, &H9D, &H66, &HF7, _
'&H18, &H1F, &H59, &H17, &H8, &H93, &H72, &HB6, &H8B, &H2A, &H13, &H95, &H81, &H43, &H78, &HD7, _
'&HBE, &H9B, &H92, &HA9, &H5A, &H65, &HDE, &H1A, &HAD, &HCE, &HB2, &H4D, &H69, &HA5, &H36, &H7F, _
'&H9F, &H46, &H53, &H33, &H47, &H21, &HD, &HFD, &H7B, &HAB, &HD1, &HF, &H84, &HFA, &HAE, &HCD, _
'&H99, &H86, &HBD, &H37, &H14, &H6B, &HDB, &HCB, &HB7, &HFB, &H3E, &H63, &HC5, &HF1, &HFE, &HC3, _
'&H82, &H5E, &H55, &H67, &H89, &HA3, &HDD, &H15, &H1B, &HD3, &HB9, &HB3, &H39, &HE9, &HEE, &H7D)
'
'sBox2 = Array(&H78, &HE, &H26, &H1D, &H80, &H77, &HD4, &H1B, &H30, &HE5, &H42, &HC2, &H24, &HD1, &HFA, &HC1, _
'&HF8, &HE4, &HDA, &H8A, &H44, &H19, &H36, &H3F, &H98, &H23, &HBC, &HCA, &H18, &H4F, &HDC, &HAC, _
'&HB8, &HC8, &H2, &H96, &H28, &H58, &H8C, &HA3, &H88, &H27, &H6C, &H89, &H34, &HDD, &H82, &H13, _
'&HE8, &H74, &HE0, &HF4, &H4C, &HD9, &H40, &H46, &HD0, &H93, &H62, &H81, &H14, &H6F, &H5A, &H64, _
'&H0, &HA4, &H66, &HE9, &HA8, &H5F, &H86, &H11, &HC, &HB, &H68, &HEE, &H1F, &HE2, &H8E, &H69, _
'&H20, &HBF, &H33, &H6B, &HF0, &HBE, &H1, &H3C, &H48, &H79, &H9C, &HD8, &H57, &HFD, &HFE, &HEA, _
'&H50, &HEC, &HBA, &H38, &H5C, &H15, &H2F, &H43, &H70, &H9B, &H22, &HA6, &H1A, &H35, &H54, &H8F, _
'&H60, &HB5, &HE6, &HC9, &H1C, &HB7, &HAA, &HD5, &H8, &HAE, &H3E, &H6A, &HCC, &HF, &H53, &HF7, _
'&H10, &H5E, &HF2, &HD, &H32, &H3A, &HCE, &HCB, &HA0, &H6, &H1E, &HFC, &H16, &H71, &H7B, &H5B, _
'&H4, &HE3, &HC4, &H9E, &H52, &H51, &H31, &H75, &H63, &H84, &H21, &HAB, &H90, &H7C, &H6E, &H12, _
'&H7A, &HB4, &HB0, &H99, &H94, &H2D, &H56, &H2E, &H39, &HB2, &HF6, &H9D, &H47, &H25, &H37, &H7, _
'&H9A, &HA, &H61, &H7D, &H4A, &H8D, &H3D, &HEB, &H17, &H91, &H49, &H29, &H2C, &H92, &HB1, &HDE, _
'&HAD, &HB3, &H3, &H7F, &HA2, &HBB, &HC6, &H55, &HA1, &HF5, &HC5, &HFB, &HC7, &HED, &H4D, &H87, _
'&HC0, &HA9, &H45, &HAF, &H2A, &H5D, &H76, &H65, &H4E, &HDF, &H41, &H9F, &HD3, &HD7, &H2B, &HC3, _
'&HA5, &H5, &H59, &H67, &H83, &HA7, &HD6, &H85, &HB6, &H73, &H3B, &HBD, &H95, &HF3, &H97, &HE1, _
'&HE7, &HF1, &H9, &H7E, &H4B, &HB9, &HDB, &HF9, &H72, &HEF, &HCD, &HD2, &H8B, &HFF, &HCF, &H6D)
'End Sub


'---------------------------------------------------------
'                      HKEYDEMO.BAS
'
'Este módulo contiene la declaración de la librería DLL con la
'función HARDkey(), las variables públicas y dos vectores que
'usan las rutinas para encriptar la cadena de comandos. Estos
'vectores son distintos para cada kit. No cambie sus valores,
'incluya este módulo en su proyecto tal como está.
'
'ATENCION: Este proyecto debe ser compilado y ejecutado
'"FUERA" del entorno de Visual Basic, ya que el mismo no
'es compatible con algunas rutinas antidebug que posee la
'dll Hkey-w32.dll
'---------------------------------------------------------
'                      HKEYDEMO.BAS
'
'This module contains declaration of DLL with
'HARDkey() function, variables publics and two vectors that
'functions uses to encript command string. These
' vectors are diferent for each development kit. Don't change it's values,
'include this module in your proyect.
'
'ATENTION: This proyect must be compiled and executed
'"OUT" of Visual Basic develop environment, because it
'is incompatible with some functions against debuggin
'in dll Hkey-w32.dll
Declare Function HARDkey Lib "hkey-w32.dll" (ByVal buf As String) As Long

'El entorno de desarrollo de Visual Basic requiere que las librerías
'DLLs se encuentren en el directorio \WINDOWS para poder abrirlas.
'El ejecutable yá compilado la encuentra tanto en el directorio \WINDOWS
'como en el directorio desde donde corre el programa.
'Visual Basic develop environment needs librarys.
'DLLs be placed in \WINDOWS directory to open.
'Compiled executable search in \WINDOWS
'and application directory.
Public sBox1 As Variant
Public sBox2 As Variant
Public password As String

Public Sub InitSB()

'Esta rutina debe ser invocada al cargar la FORM donde se realiza el
'chequeo de la llave. En ella se inicializan los vectores sBox que
'usan las rutinas EncriptaString() y DesencriptString() para encriptar
'la cadena que se pasa a la funcion HARDkey().

'This procedure must be call before checking key
'It initialize sBox vector for EncriptaString() and DesencriptString()
'procedures.
password = "DVXQXRSvdskQZAvc"

    'Debe coincidir con la password que declaró en el campo "PASSWORD
    'para encriptado de parámetros" del instalador del kit.
    'Debe tener una longitud fija de 16 caracteres.

    'It must be password you declare when install kit
    'The length can't exceed 16 characters.

sBox1 = Array(&H9E, &H70, &H5C, &HCE, &HA6, &H80, &H66, &HE0, &H1E, &HF4, &HE6, &H25, &H8, &H6B, &H74, &HCF, _
&H0, &H8B, &HEA, &HD9, &HB2, &H13, &HBE, &HB5, &HDE, &HFD, &HEC, &H73, &H82, &H2D, &HC2, &H97, _
&H16, &HA3, &H3, &H5A, &HF2, &HD7, &HD8, &H4A, &HF6, &HDA, &H50, &H52, &H7A, &H3B, &H58, &H15, _
&H3E, &HF0, &H4C, &H17, &H12, &HD5, &HDC, &H29, &HD6, &H34, &HBC, &H6E, &H1C, &HFA, &H3C, &HD3, _
&H10, &HF7, &H68, &HB6, &H76, &H37, &HA0, &HED, &H40, &HEE, &H7, &HB0, &H20, &H45, &H8C, &H4F, _
&H8E, &H3A, &H2, &H4E, &H9A, &HBA, &HE8, &H3F, &H2A, &HB4, &HFC, &H61, &H49, &H6A, &HE, &H1D, _
&HFE, &HE3, &H98, &H4D, &H6, &H92, &H28, &HC1, &H44, &H95, &H4, &H63, &H24, &HC4, &H86, &HCB, _
&H7E, &HB8, &HD4, &H6C, &HA2, &HA9, &H31, &H30, &H22, &HCD, &H39, &HF3, &H32, &HD1, &HF8, &H19, _
&H48, &H18, &HAC, &HEB, &H96, &HCC, &H54, &H5B, &H84, &HD0, &H5D, &H9D, &H2C, &H3D, &H5, &H36, _
&H81, &H53, &H51, &H8F, &HE2, &H14, &H5E, &H79, &H60, &H93, &HAE, &H9C, &HC6, &H65, &H8D, &H23, _
&H7C, &HE1, &HCA, &H7F, &H2B, &HC9, &H8A, &HB1, &H43, &HC0, &H38, &HB3, &H2E, &H69, &H77, &H91, _
&H46, &H35, &H56, &H85, &H4B, &H6F, &H55, &HF1, &H9B, &HC8, &H59, &HE4, &H1B, &H94, &H9F, &HA1, _
&H90, &HD2, &H78, &H2F, &H62, &H7D, &HAA, &H1F, &H71, &HD, &HAB, &HC5, &HA5, &HAF, &HC7, &H83, _
&HA4, &HDD, &H1, &H33, &HBD, &HE7, &HBB, &H6D, &HB, &H42, &H75, &H11, &H89, &HA7, &HAD, &HDB, _
&H26, &HB7, &H1A, &H41, &HDF, &HE9, &HA8, &H21, &H64, &HF, &HA, &HFB, &H27, &H5F, &H99, &H9, _
&H72, &HC, &HC3, &HF9, &HB9, &H47, &HF5, &H57, &HEF, &H67, &HBF, &H7B, &HE5, &HFF, &H88, &H87)

sBox2 = Array(&H1E, &HDF, &HB4, &H96, &H6, &HD7, &H90, &HB9, &H36, &H78, &H72, &H2, &H8, &H52, &H9A, &HB5, _
&HCE, &HCB, &HE, &H8E, &H3A, &HF6, &HCC, &H83, &HC, &H7E, &HEE, &H69, &H6A, &H40, &H84, &HAF, _
&H5E, &H20, &H4E, &HFE, &H4A, &HC4, &HD8, &HAA, &HC6, &HB1, &H7C, &H16, &HDA, &HD, &HEA, &H2B, _
&HDE, &H8D, &H15, &H26, &HA6, &H63, &HBA, &HF8, &HE6, &H66, &H2F, &H47, &HB6, &HAC, &H10, &H19, _
&H9E, &H7D, &H1A, &H77, &HB2, &H2A, &H68, &H5, &H22, &H82, &H3D, &H98, &H3F, &H7F, &HF0, &HEF, _
&H3E, &HFC, &H94, &H79, &HE2, &HB8, &H4, &H48, &H23, &H13, &H2E, &H32, &H50, &HFF, &H58, &H86, _
&HBE, &HE0, &HA0, &HCD, &HD2, &H35, &H3C, &HDD, &H30, &H33, &HAE, &H5A, &H27, &HF1, &H18, &HBC, _
&H59, &H95, &H6C, &HC7, &H2C, &HE3, &H4B, &H64, &H70, &H5D, &HA8, &H1D, &H4C, &HFD, &H12, &HDC, _
&H7B, &H2D, &H80, &HA1, &HF2, &H24, &H3, &H73, &H28, &HE7, &H85, &HEC, &H88, &H71, &H92, &HE1, _
&H60, &H6B, &H4D, &H53, &H6E, &HB, &H1B, &H9, &H34, &HE8, &H51, &H9F, &H46, &HC2, &H38, &HD9, _
&H21, &HAD, &H43, &HCF, &H39, &H44, &H97, &H62, &H42, &HB0, &H31, &H91, &H74, &HD6, &H5F, &H3B, _
&H5C, &HA2, &H99, &HE9, &H49, &HDB, &H25, &H14, &H5B, &HB3, &H57, &HE4, &H0, &HA9, &H81, &HD3, _
&H9C, &HA5, &H9D, &HBD, &HC3, &HC5, &H8C, &H67, &HA4, &HC9, &H1C, &HED, &H54, &H8F, &H7A, &HC1, _
&HC8, &HBB, &H8B, &H65, &H6F, &HF3, &HA, &H8A, &H55, &HD1, &H37, &HA7, &HD0, &H11, &H17, &H87, _
&HF, &HD5, &HD4, &H41, &HB7, &HE5, &H56, &HF7, &H76, &HFB, &H9B, &H6D, &HA3, &H61, &H1F, &HCA, _
&H1, &H75, &HC0, &HFA, &HBF, &HEB, &HF4, &H93, &HAB, &H29, &H45, &HF5, &HF9, &H4F, &H7, &H89)
End Sub


