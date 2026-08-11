VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum contrato
    PI_Saldo = 0
    PI_ChequeDif = 1
    PI_Letras = 2
    PI_PostVenta = 3
End Enum

Public Enum Operacion
    IniciarConexion = 0
    VerificarConexion = 1
    FinalizarConexion = 2
    LeerCelda = 3
    GrabarCelda = 4
    LeerCadena = 5
    GrabarCadena = 6
    LeerRTC = 10
    GrabarRTC = 11
End Enum

Public Enum Celda
    Nro00 = 0
    Nro01 = 1
    Nro02 = 2
    Nro03 = 3
    Nro04 = 4
    Nro05 = 5
    Nro06 = 6
    Nro07 = 7
    Nro08 = 8
    Nro09 = 9
    Nro10 = 10
    Nro11 = 11
    Nro12 = 12
    Nro13 = 13
    Nro14 = 14
    Nro15 = 15
    Nro16 = 16
    Nro17 = 17
    Nro18 = 18
    Nro19 = 19
    Nro20 = 20
    Nro21 = 21
    Nro22 = 22
    Nro23 = 23
    Nro24 = 24
    Nro25 = 25
    Nro26 = 26
    Nro27 = 27
    Nro28 = 28
    Nro29 = 29
    Nro30 = 30
    Nro31 = 31
    Nro32 = 32
    Nro33 = 33
    Nro34 = 34
    Nro35 = 35
    Nro36 = 36
    Nro37 = 37
    Nro38 = 38
    Nro39 = 39
    Nro40 = 40
    Nro41 = 41
    Nro42 = 42
    Nro43 = 43
    Nro44 = 44
    Nro45 = 45
    Nro46 = 46
    Nro47 = 47
    Nro48 = 48
    Nro49 = 49
    Nro50 = 50
    Nro51 = 51
    Nro52 = 52
    Nro53 = 53
End Enum

Public Enum Bite
    Byte_Nro44 = 44
    Byte_Nro45 = 45
    Byte_Nro46 = 46
    Byte_Nro47 = 47
    Byte_Nro48 = 48
    Byte_Nro49 = 49
    Byte_Nro50 = 50
    Byte_Nro51 = 51
    Byte_Nro52 = 52
    Byte_Nro53 = 53
    Byte_Nro54 = 54
    Byte_Nro55 = 55
    Byte_Nro56 = 56
    Byte_Nro57 = 57
    Byte_Nro58 = 58
    Byte_Nro59 = 59
    Byte_Nro60 = 60
    Byte_Nro61 = 61
    Byte_Nro62 = 62
    Byte_Nro63 = 63
    Byte_Nro64 = 64
    Byte_Nro65 = 65
    Byte_Nro66 = 66
    Byte_Nro67 = 67
    Byte_Nro68 = 68
    Byte_Nro69 = 69
    Byte_Nro70 = 70
    Byte_Nro71 = 71
    Byte_Nro72 = 72
    Byte_Nro73 = 73
    Byte_Nro74 = 74
    Byte_Nro75 = 75
    Byte_Nro76 = 76
    Byte_Nro77 = 77
    Byte_Nro78 = 78
    Byte_Nro79 = 79
    Byte_Nro80 = 80
    Byte_Nro81 = 81
    Byte_Nro82 = 82
    Byte_Nro83 = 83
    Byte_Nro84 = 84
    Byte_Nro85 = 85
    Byte_Nro86 = 86
    Byte_Nro87 = 87
    Byte_Nro88 = 88
    Byte_Nro89 = 89
    Byte_Nro90 = 90
    Byte_Nro91 = 91
    Byte_Nro92 = 92
    Byte_Nro93 = 93
    Byte_Nro94 = 94
    Byte_Nro95 = 95
    Byte_Nro96 = 96
    Byte_Nro97 = 97
    Byte_Nro98 = 98
    Byte_Nro99 = 99
    Byte_Nro100 = 100
    Byte_Nro101 = 101
    Byte_Nro102 = 102
    Byte_Nro103 = 103
    Byte_Nro104 = 104
    Byte_Nro105 = 105
    Byte_Nro106 = 106
    Byte_Nro107 = 107
End Enum

Public Enum Aplicacion
    Adicion = 3 'N
    Administracion = 9 'A
    PuntoVenta = 13 'P
    Consultas = 25 'C
    CentralDelivery = 17 'D
    
    Almacen = 19 'L
    Costos = 5 'S
    
    Infhotel = 10 'I
    Eventos = 22 'E
    Despachador = 11 'H
    Motorizado = 2 'M
    Telefonia = 8 'F
    MotorConta = 21 'T
    CheffControl = 26 'R
    
End Enum

