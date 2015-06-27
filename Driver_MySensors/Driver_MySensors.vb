Imports HoMIDom
Imports HoMIDom.HoMIDom.Device
Imports HoMIDom.HoMIDom.Server
Imports STRGS = Microsoft.VisualBasic.Strings
Imports System.IO.Ports
Imports System.Text.RegularExpressions

'************************************************
'INFOS 
'************************************************
'Le driver communique en "COM" avec l'arduino gateway qui doit implémenter un sketch spécifique compatible MySensors
'http://mysensors.org
'************************************************

Public Class Driver_Arduino_USB
    Implements HoMIDom.HoMIDom.IDriver

#Region "Variables génériques"
    '!!!Attention les variables ci-dessous doivent avoir une valeur par défaut obligatoirement
    'aller sur l'adresse http://www.somacon.com/p113.php pour avoir un ID
    Dim _ID As String = "A5B6D4C4-FF24-11E4-9733-BF931D5D46B0"
    Dim _Nom As String = "MySensors"
    Dim _Enable As Boolean = False
    Dim _Description As String = "Driver MySensors"
    Dim _StartAuto As Boolean = False
    Dim _Protocol As String = "COM"
    Dim _IsConnect As Boolean = False
    Dim _IP_TCP As String = "@"
    Dim _Port_TCP As String = "@"
    Dim _IP_UDP As String = "@"
    Dim _Port_UDP As String = "@"
    Dim _Com As String = ""
    Dim _Refresh As Integer = 0
    Dim _Modele As String = "@"
    Dim _Version As String = My.Application.Info.Version.ToString
    Dim _OsPlatform As String = "3264"
    Dim _Picture As String = ""
    Dim _Server As HoMIDom.HoMIDom.Server
    Dim _Device As HoMIDom.HoMIDom.Device
    Dim _DeviceSupport As New ArrayList
    Dim _Parametres As New ArrayList
    Dim _LabelsDriver As New ArrayList
    Dim _LabelsDevice As New ArrayList
    Dim MyTimer As New Timers.Timer
    Dim _idsrv As String
    Dim _DeviceCommandPlus As New List(Of HoMIDom.HoMIDom.Device.DeviceCommande)
    Dim _AutoDiscover As Boolean = False
    Dim _acknowledge As Boolean = False

    'param avancé
    Dim _DEBUG As Boolean = False

#End Region

#Region "Variables Internes"
    Private serialPortObj As SerialPort
    'Public WithEvents port As New System.IO.Ports.SerialPort
    Dim _BAUD As Integer = 9600
    Dim _RCVERROR As Boolean = True
    Dim first As Boolean = False
#End Region

#Region "Propriétés génériques"
    Public WriteOnly Property IdSrv As String Implements HoMIDom.HoMIDom.IDriver.IdSrv
        Set(ByVal value As String)
            _idsrv = value
        End Set
    End Property

    Public Property COM() As String Implements HoMIDom.HoMIDom.IDriver.COM
        Get
            Return _Com
        End Get
        Set(ByVal value As String)
            _Com = value
        End Set
    End Property
    Public ReadOnly Property Description() As String Implements HoMIDom.HoMIDom.IDriver.Description
        Get
            Return _Description
        End Get
    End Property
    Public ReadOnly Property DeviceSupport() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.DeviceSupport
        Get
            Return _DeviceSupport
        End Get
    End Property

    Public Property Parametres() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.Parametres
        Get
            Return _Parametres
        End Get
        Set(ByVal value As System.Collections.ArrayList)
            _Parametres = value
        End Set
    End Property

    Public Property LabelsDriver() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.LabelsDriver
        Get
            Return _LabelsDriver
        End Get
        Set(ByVal value As System.Collections.ArrayList)
            _LabelsDriver = value
        End Set
    End Property
    Public Property LabelsDevice() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.LabelsDevice
        Get
            Return _LabelsDevice
        End Get
        Set(ByVal value As System.Collections.ArrayList)
            _LabelsDevice = value
        End Set
    End Property

    Public Event DriverEvent(ByVal DriveName As String, ByVal TypeEvent As String, ByVal Parametre As Object) Implements HoMIDom.HoMIDom.IDriver.DriverEvent

    Public Property Enable() As Boolean Implements HoMIDom.HoMIDom.IDriver.Enable
        Get
            Return _Enable
        End Get
        Set(ByVal value As Boolean)
            _Enable = value
        End Set
    End Property
    Public ReadOnly Property ID() As String Implements HoMIDom.HoMIDom.IDriver.ID
        Get
            Return _ID
        End Get
    End Property
    Public Property IP_TCP() As String Implements HoMIDom.HoMIDom.IDriver.IP_TCP
        Get
            Return _IP_TCP
        End Get
        Set(ByVal value As String)
            _IP_TCP = value
        End Set
    End Property
    Public Property IP_UDP() As String Implements HoMIDom.HoMIDom.IDriver.IP_UDP
        Get
            Return _IP_UDP
        End Get
        Set(ByVal value As String)
            _IP_UDP = value
        End Set
    End Property
    Public ReadOnly Property IsConnect() As Boolean Implements HoMIDom.HoMIDom.IDriver.IsConnect
        Get
            Return _IsConnect
        End Get
    End Property
    Public Property Modele() As String Implements HoMIDom.HoMIDom.IDriver.Modele
        Get
            Return _Modele
        End Get
        Set(ByVal value As String)
            _Modele = value
        End Set
    End Property
    Public ReadOnly Property Nom() As String Implements HoMIDom.HoMIDom.IDriver.Nom
        Get
            Return _Nom
        End Get
    End Property
    Public Property Picture() As String Implements HoMIDom.HoMIDom.IDriver.Picture
        Get
            Return _Picture
        End Get
        Set(ByVal value As String)
            _Picture = value
        End Set
    End Property
    Public Property Port_TCP() As String Implements HoMIDom.HoMIDom.IDriver.Port_TCP
        Get
            Return _Port_TCP
        End Get
        Set(ByVal value As String)
            _Port_TCP = value
        End Set
    End Property
    Public Property Port_UDP() As String Implements HoMIDom.HoMIDom.IDriver.Port_UDP
        Get
            Return _Port_UDP
        End Get
        Set(ByVal value As String)
            _Port_UDP = value
        End Set
    End Property
    Public ReadOnly Property Protocol() As String Implements HoMIDom.HoMIDom.IDriver.Protocol
        Get
            Return _Protocol
        End Get
    End Property
    Public Property Refresh() As Integer Implements HoMIDom.HoMIDom.IDriver.Refresh
        Get
            Return _Refresh
        End Get
        Set(ByVal value As Integer)
            _Refresh = value
        End Set
    End Property
    Public Property Server() As HoMIDom.HoMIDom.Server Implements HoMIDom.HoMIDom.IDriver.Server
        Get
            Return _Server
        End Get
        Set(ByVal value As HoMIDom.HoMIDom.Server)
            _Server = value
        End Set
    End Property
    Public ReadOnly Property Version() As String Implements HoMIDom.HoMIDom.IDriver.Version
        Get
            Return _Version
        End Get
    End Property
    Public ReadOnly Property OsPlatform() As String Implements HoMIDom.HoMIDom.IDriver.OsPlatform
        Get
            Return _OsPlatform
        End Get
    End Property
    Public Property StartAuto() As Boolean Implements HoMIDom.HoMIDom.IDriver.StartAuto
        Get
            Return _StartAuto
        End Get
        Set(ByVal value As Boolean)
            _StartAuto = value
        End Set
    End Property
    Public Property AutoDiscover() As Boolean Implements HoMIDom.HoMIDom.IDriver.AutoDiscover
        Get
            Return _AutoDiscover
        End Get
        Set(ByVal value As Boolean)
            _AutoDiscover = value
        End Set
    End Property
#End Region

#Region "Fonctions génériques"
    ''' <summary>
    ''' Retourne la liste des Commandes avancées
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCommandPlus() As List(Of DeviceCommande)
        Return _DeviceCommandPlus
    End Function

    ''' <summary>Execute une commande avancée</summary>
    ''' <param name="MyDevice">Objet représentant le Device </param>
    ''' <param name="Command">Nom de la commande avancée à éxécuter</param>
    ''' <param name="Param">tableau de paramétres</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteCommand(ByVal MyDevice As Object, ByVal Command As String, Optional ByVal Param() As Object = Nothing) As Boolean
        Dim retour As Boolean = False
        Try
            If MyDevice IsNot Nothing Then
                'Pas de commande demandée donc erreur
                If Command = "" Then
                    Return False
                Else
                    Write(MyDevice, Command, Param(0), Param(1))
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " ExecuteCommand", "exception : " & ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Permet de vérifier si un champ est valide
    ''' </summary>
    ''' <param name="Champ"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function VerifChamp(ByVal Champ As String, ByVal Value As Object) As String Implements HoMIDom.HoMIDom.IDriver.VerifChamp
        Try
            Dim retour As String = "0"
            Select Case UCase(Champ)
                Case "ADRESSE1"
                    If Value = " " Then retour = "l'adresse est obligatoire"
                Case "ADRESSE2"
                    If Value = " " Then retour = "l'adresse est obligatoire"
            End Select
            Return retour
        Catch ex As Exception
            Return "Une erreur est apparue lors de la vérification du champ " & Champ & ": " & ex.ToString
        End Try
    End Function

    ''' <summary>Démarrer le du driver</summary>
    ''' <remarks></remarks>
    Public Sub Start() Implements HoMIDom.HoMIDom.IDriver.Start
        Try
            If Not _IsConnect Then
                Dim trv As Boolean = False
                Dim _ports As String = "<AUCUN>"

                'récupération des paramétres avancés
                Try
                    _DEBUG = _Parametres.Item(0).Valeur
                    _BAUD = _Parametres.Item(1).Valeur
                    _RCVERROR = _Parametres.Item(2).Valeur
                Catch ex As Exception
                    _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Start", "ERR: Erreur dans les paramétres avancés. utilisation des valeur par défaut" & ex.Message)
                    _DEBUG = False
                    _BAUD = 57600
                    _RCVERROR = True
                End Try

                If _Com = "" Or _Com = " " Then
                    _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Start", "Le port COM est vide veuillez le renseigner")
                    Exit Sub
                End If

                Dim portNames As String() = SerialPort.GetPortNames()
                Array.Sort(portNames)
                For Each serialPortName As String In portNames
                    _ports &= serialPortName & " "
                    If UCase(serialPortName) = UCase(_Com) Then
                        trv = True
                    End If
                Next

                If trv = False Then
                    _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Start", "Le port COM " & _Com & " n'existe pas, seuls les ports " & _ports & " existe(s)!")
                    Exit Sub
                End If

                serialPortObj = New SerialPort()
                serialPortObj.PortName = _Com
                serialPortObj.BaudRate = _BAUD
                serialPortObj.Parity = Parity.None
                serialPortObj.DataBits = 8
                serialPortObj.StopBits = 1
                serialPortObj.ReadTimeout = 50000
                serialPortObj.Encoding = System.Text.Encoding.GetEncoding("ISO-8859-1")

                If _RCVERROR Then AddHandler serialPortObj.ErrorReceived, New SerialErrorReceivedEventHandler(AddressOf serialPortObj_ErrorReceived)
                AddHandler serialPortObj.DataReceived, New SerialDataReceivedEventHandler(AddressOf DataReceived)

                If serialPortObj.IsOpen Then
                    serialPortObj.Close()
                End If

                serialPortObj.Open()
                serialPortObj.DiscardInBuffer()
                _IsConnect = True
                first = True
                _Server.Log(TypeLog.INFO, TypeSource.DRIVER, Me.Nom & " Start", "Port " & _Com & " ouvert")
            Else
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Start", "Port " & _Com & " déjà ouvert")
            End If

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Start", ex.ToString)
        End Try
    End Sub

    ''' <summary>Arrêter le du driver</summary>
    ''' <remarks></remarks>
    Public Sub [Stop]() Implements HoMIDom.HoMIDom.IDriver.Stop
        Try
            If _IsConnect Then
                serialPortObj.Close()
                RemoveHandler serialPortObj.ErrorReceived, New SerialErrorReceivedEventHandler(AddressOf serialPortObj_ErrorReceived)
                RemoveHandler serialPortObj.DataReceived, New SerialDataReceivedEventHandler(AddressOf DataReceived)
                _IsConnect = False
            End If
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Stop", ex.Message)
        End Try
    End Sub

    ''' <summary>Re-Démarrer le du driver</summary>
    ''' <remarks></remarks>
    Public Sub Restart() Implements HoMIDom.HoMIDom.IDriver.Restart
        [Stop]()
        Start()
    End Sub

    ''' <summary>Intérroger un device</summary>
    ''' <param name="Objet">Objet représetant le device à interroger</param>
    ''' <remarks></remarks>
    Public Sub Read(ByVal Objet As Object) Implements HoMIDom.HoMIDom.IDriver.Read
        Try
            If _Enable = False Then Exit Sub
            If _IsConnect = False Then
                WriteLog("Le driver n'est pas démarré, impossible de communiquer avec l'arduino")
                Exit Sub
            End If
            If _DEBUG Then WriteLog("DBG: WRITE Device " & Objet.Name & " <-- " & Command())

            'verification si adresse1 n'est pas vide
            If String.IsNullOrEmpty(Objet.Adresse1) Or Objet.Adresse1 = "" Then
                WriteLog("ERR: WRITE l'adresse de l'arduino doit etre renseigné (ex: 1 pour un arduino maitre, 2-255 pour un arduino esclave) : " & Objet.Name)
                Exit Sub
            End If


            'suivant le type du PIN on lance la bonne commande : ENTREE_ANA|ENTREE_DIG|SORTIE_DIG|PWM|1WIRE

            Dim MySensorsCommand As String = ""
            Select Case UCase(Objet.Modele)
                Case "VARIABLE"
                    MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " VR " & Objet.Adresse2
                    WriteLog("DBG: Commande passée à l arduino de type : VARIABLE")
                Case "ANALOG_IN"
                    MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " AR " & Objet.Adresse2
                    WriteLog("DBG: Commande passée à l arduino de type : ANALOG_IN")
                Case "DIGITAL_IN"
                    WriteLog("DBG: Commande passée à l arduino de type : DIGITAL_IN")
                    MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " DR " & Objet.Adresse2
                Case "DHTXX"
                    WriteLog("DBG: Commande passée à l arduino de type : DHTXX")
                    MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " DHTXX " & Objet.Adresse2
                Case "BMP180"
                    WriteLog("DBG: Commande passée à l arduino de type : BMP180")
                    MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " BMP180 " & Objet.Adresse2
                Case "CUSTOM"
                    WriteLog("DBG: Commande passée à l arduino de type : CUSTOM")
                    MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " " & Objet.Adresse2
                Case "1WIRE"
                    WriteLog("le 1-wire n'est pas encore géré :" & Objet.Name)
                    Exit Sub
                Case ""
                    WriteLog("ERR: WRITE Pas de protocole d'emission pour " & Objet.Name)
                    Exit Sub
                Case Else
                    WriteLog("ERR: WRITE Protocole non géré : " & Objet.Modele.ToString.ToUpper)
                    Exit Sub
            End Select
            'End If

            'If MySensorsCommand <> "" Then
            'If _DEBUG Then WriteLog("DBG: WRITE Composant " & Objet.Name & " URL : " & MySensorsCommand)

            WriteLog("DBG: Commande passée à l arduino : " & MySensorsCommand)
            serialPortObj.WriteLine(MySensorsCommand) ', 0, 8)
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Read", ex.Message)
        End Try
    End Sub

    ''' <summary>Commander un device</summary>
    ''' <param name="Objet">Objet représetant le device à interroger</param>
    ''' <param name="Command">La commande à passer</param>
    ''' <param name="Parametre1"></param>
    ''' <param name="Parametre2"></param>
    ''' <remarks></remarks>
    Public Sub Write(ByVal Objet As Object, ByVal Command As String, Optional ByVal Parametre1 As Object = Nothing, Optional ByVal Parametre2 As Object = Nothing) Implements HoMIDom.HoMIDom.IDriver.Write
        Try
            If _Enable = False Then Exit Sub
            If _IsConnect = False Then
                WriteLog("Le driver n'est pas démarré, impossible de communiquer avec l'arduino")
                Exit Sub
            End If
            If _DEBUG Then WriteLog("DBG: WRITE Device " & Objet.Name & " <-- " & Command)

            'verification si adresse1 n'est pas vide
            If String.IsNullOrEmpty(Objet.Adresse1) Or Objet.Adresse1 = "" Then
                WriteLog("ERR: WRITE l'adresse de l'arduino doit etre renseigné (ex: 1 pour un arduino maitre, 2-255 pour un arduino esclave) : " & Objet.Name)
                Exit Sub
            End If

            Dim MySensorsCommand As String = ""
            Select Case UCase(Objet.Modele)
                Case "V_LIGHT"
                    Select Case Command
                        Case "ON"
                            MySensorsCommand = Objet.Adresse1 & ";" & Objet.adresse2 & ";1;1;2;1"
                        Case "OFF"
                            MySensorsCommand = Objet.Adresse1 & ";" & Objet.adresse2 & ";1;1;2;0"
                    End Select
                Case "V_DIMMER"
                    Select Case Command
                        Case "ON"
                            MySensorsCommand = Objet.Adresse1 & ";" & Objet.adresse2 & ";1;1;3;100"
                        Case "OFF"
                            MySensorsCommand = Objet.Adresse1 & ";" & Objet.adresse2 & ";1;1;3;0"
                        Case "DIM"
                            If Not IsNothing(Parametre1) Then
                                If IsNumeric(Parametre1) Then
                                    ''Conversion du parametre de % (0 à 100) en 0 à 255
                                    'Parametre1 = CInt(Parametre1 * 255 / 100)
                                    MySensorsCommand = Objet.Adresse1 & ";" & Objet.adresse2 & ";1;1;3;" & Parametre1
                                Else
                                    WriteLog("ERR: WRITE DIM Le parametre " & CStr(Parametre1) & " n'est pas un entier (" & Objet.Name & ")")
                                End If
                            Else
                                WriteLog("ERR: WRITE DIM Il manque un parametre (" & Objet.Name & ")")
                            End If
                        Case "PWM"
                            If Not IsNothing(Parametre1) Then
                                If IsNumeric(Parametre1) Then
                                    If CInt(Parametre1) > 255 Then Parametre1 = 255
                                    If CInt(Parametre1) < 0 Then Parametre1 = 0
                                    'Conversion du parametre de 0 à 255 en % (0 à 100)
                                    Parametre1 = CInt(Parametre1 * 100 / 255)
                                    MySensorsCommand = Objet.Adresse1 & ";" & Objet.adresse2 & ";1;1;3;" & Parametre1
                                Else
                                    WriteLog("ERR: WRITE DIM Le parametre " & CStr(Parametre1) & " n'est pas un entier (" & Objet.Name & ")")
                                End If
                            Else
                                WriteLog("ERR: WRITE DIM Il manque un parametre (" & Objet.Name & ")")
                            End If
                        Case Else
                            WriteLog("ERR: Send AC : Commande invalide : " & Command & " (ON/OFF/DIM/PWM supporté sur une SORTIE: Analogique write)")
                            Exit Sub
                    End Select
                Case Else
                    WriteLog("ERR: WRITE : Ce type de Sensor/Pin ne peut pas être piloté : " & Objet.Modele.ToString.ToUpper & " (" & Objet.Name & ")")
                    Exit Sub
            End Select
            'If Command = "CONFIG_TYPE_PIN" Then
            '    Select Case UCase(Objet.Modele)
            '        Case "ANALOG_IN"
            '            MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " PM " & Objet.Adresse2 & " 0"
            '        Case "DIGITAL_IN"
            '            MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " PM " & Objet.Adresse2 & " 0"
            '        Case "DIGITAL_OUT"
            '            MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " PM " & Objet.Adresse2 & " 1"
            '        Case "DIGITAL_PWM"
            '            MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " PM " & Objet.Adresse2 & " 1"
            '        Case "LCD"
            '            MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " PM " & Objet.Adresse2 & " 1"
            '        Case "KPAD"
            '            ' Pas d'initialisation
            '        Case "1WIRE"
            '            ' Pas d'initialisation
            '        Case "DHTxx"
            '            ' Pas d'initialisation
            '        Case "BMP180"
            '            ' Pas d'initialisation
            '        Case Else
            '            WriteLog("ERR: WRITE CONFIG_TYPE_PIN : Ce type de PIN ne peut pas être configuré : " & Objet.Modele.ToString.ToUpper & " (" & Objet.Name & ")")
            '            Exit Sub
            '    End Select
            'ElseIf Command = "SETLCD" Then
            '    WriteLog("DBG: Commande passée à l arduino de type : SETLCD")
            '    Select Case UCase(Objet.Modele)
            '        Case "LCD"
            '            If Not IsNothing(Parametre1) And Not IsNothing(Parametre2) Then
            '                Dim Parametre2a As String = Parametre2.Replace(" ", "/#")
            '                MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " SETLCD " & Parametre1 & " " & Parametre2a
            '                WriteLog("DBG: Commande passée à l arduino : " & MySensorsCommand)
            '            Else
            '                WriteLog("ERR: WRITE DIM Il manque au moins un parametre (" & Objet.Name & ")")
            '            End If
            '        Case Else
            '            WriteLog("ERR: Send AC : Commande invalide : " & Command & " (ON/OFF ou SETLCD supportés)")
            '            Exit Sub
            '    End Select
            '    'Else
            '    '   WriteLog("ERR: WRITE SETVAR : Il manque la valeur à passer à la variable en parametre : " & Objet.Modele.ToString.ToUpper & " (" & Objet.Name & ")")
            '    '  Exit Sub
            '    'End If
            'ElseIf Command = "SETVAR" Then
            '    If Not IsNothing(Parametre1) Then
            '        Select Case UCase(Objet.Modele)
            '            Case "VARIABLE" : MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " VW " & Objet.Adresse2 & "_" & Parametre1
            '            Case Else
            '                WriteLog("ERR: WRITE SETVAR : Seulement une variable peut utilisé la fonction SETVAR : " & Objet.Modele.ToString.ToUpper & " (" & Objet.Name & ")")
            '                Exit Sub
            '        End Select
            '    Else
            '        WriteLog("ERR: WRITE SETVAR : Il manque la valeur à passer à la variable en parametre : " & Objet.Modele.ToString.ToUpper & " (" & Objet.Name & ")")
            '        Exit Sub
            '    End If
            'ElseIf Command = "READX" Then
            '    '    MySensorsCommand = "ACT ADR " & Objet.Adresse1 & "/?homidom_READX"
            'Else
            '    Select Case UCase(Objet.Modele)
            '        'Case "ARMED"
            '        'Case "CURRENT"
            '        'Case "DIMMER"
            '        'Case "DIRECTION"
            '        'Case "DISTANCE"
            '        'Case "DOWN"
            '        'Case "DUST_LEVEL"
            '        'Case "FLOWVOLUME"
            '        'Case "FORECAST"
            '        'Case "GUST"
            '        'Case "HEATER"
            '        'Case "HEATER_SW"
            '        'Case "HUM"
            '        'Case "IMPEDANCE"
            '        'Case "IR_SEND"
            '        'Case "IR_RECEIVE"
            '        'Case "LIGHT"
            '        'Case "LIGHT_LEVEL"
            '        'Case "LOCK_STATUS"
            '        'Case "KWH"
            '        'Case "PRESSURE"
            '        'Case "RAIN"
            '        'Case "RAINRATE"
            '        'Case "SCENE_ON"
            '        'Case "SCENE_OFF"
            '        'Case "TEMP"
            '        'Case "STOP"
            '        'Case "TRIPPED"
            '        'Case "UP"
            '        'Case "UV"
            '        'Case "VAR1"
            '        'Case "VAR2"
            '        'Case "VAR3"
            '        'Case "VAR4"
            '        'Case "VAR5"
            '        'Case "VOLTAGE"
            '        'Case "WATT"
            '        'Case "WEIGHT"
            '        'Case "WIND"
            '        Case "CUSTOM"
            '            MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " " & Objet.Adresse2
            '            WriteLog("DBG: Commande passée à l arduino de type : ANALOG_IN")
            '        Case "ANALOG_IN"
            '            MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " AR " & Objet.Adresse2
            '            WriteLog("DBG: Commande passée à l arduino de type : ANALOG_IN")
            '        Case "DIGITAL_IN"
            '            WriteLog("DBG: Commande passée à l arduino de type : DIGITAL_IN")
            '            MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " DR " & Objet.Adresse2
            '        Case "DIGITAL_OUT"
            '            'Digital Write
            '            WriteLog("DBG: Commande passée à l arduino de type : DIGITAL_OUT")
            '            Select Case Command
            '                Case "ON" : MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " DW " & Objet.Adresse2 & " 1"
            '                Case "OFF" : MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " DW " & Objet.Adresse2 & " 0"
            '                Case Else
            '                    WriteLog("ERR: Send AC : Commande invalide : " & Command & " (ON/OFF supporté sur une SORTIE: digital write)")
            '                    Exit Sub
            '            End Select
            '        Case "DIGITAL_PWM"
            '            'Analogique write (0-255)
            '            'on convertit ON/OFF/DIM en DIM de 0 à 255 (commande PWM sur l'arduino)
            '            WriteLog("DBG: Commande passée à l arduino de type : DIGITAL_PWM")
            '            Select Case Command
            '                Case "ON" : MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " PWM " & Objet.Adresse2 & " 255"
            '                Case "OFF" : MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " PWM " & Objet.Adresse2 & " 0"
            '                Case "DIM"
            '                    If Not IsNothing(Parametre1) Then
            '                        If IsNumeric(Parametre1) Then
            '                            'Conversion du parametre de % (0 à 100) en 0 à 255
            '                            Parametre1 = CInt(Parametre1 * 255 / 100)
            '                            MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " DW " & Objet.Adresse2 & " " & Parametre1
            '                        Else
            '                            WriteLog("ERR: WRITE DIM Le parametre " & CStr(Parametre1) & " n'est pas un entier (" & Objet.Name & ")")
            '                        End If
            '                    Else
            '                        WriteLog("ERR: WRITE DIM Il manque un parametre (" & Objet.Name & ")")
            '                    End If
            '                Case "PWM"
            '                    If Not IsNothing(Parametre1) Then
            '                        If IsNumeric(Parametre1) Then
            '                            If CInt(Parametre1) > 255 Then Parametre1 = 255
            '                            If CInt(Parametre1) < 0 Then Parametre1 = 0
            '                            MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " PWM " & Objet.Adresse1 & " " & Parametre1
            '                        Else
            '                            WriteLog("ERR: WRITE DIM Le parametre " & CStr(Parametre1) & " n'est pas un entier (" & Objet.Name & ")")
            '                        End If
            '                    Else
            '                        WriteLog("ERR: WRITE DIM Il manque un parametre (" & Objet.Name & ")")
            '                    End If
            '                Case Else
            '                    WriteLog("ERR: Send AC : Commande invalide : " & Command & " (ON/OFF/DIM/PWM supporté sur une SORTIE: Analogique write)")
            '                    Exit Sub
            '            End Select
            '        Case "LCD"
            '            'Digital Write
            '            WriteLog("DBG: Commande passée à l arduino de type : LCD")
            '            Select Case Command
            '                Case "ON" : MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " LCD " & Objet.Adresse2 & " 1"
            '                Case "OFF" : MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " LCD " & Objet.Adresse2 & " 0"
            '                    'Case "SETLCD"
            '                    '    If Not IsNothing(Parametre1) And Not IsNothing(Parametre2) Then
            '                    '        MySensorsCommand = "ACT ADR " & Objet.Adresse1 & " SETLCD " & Parametre1 & " " & Parametre2
            '                    '    Else
            '                    '        WriteLog("ERR: WRITE DIM Il manque au moins un parametre (" & Objet.Name & ")")
            '                    '    End If
            '                Case Else
            '                    WriteLog("ERR: Send AC : Commande invalide : " & Command & " (ON/OFF ou SETLCD supportés)")
            '                    Exit Sub
            '            End Select
            '        Case ""
            '            WriteLog("ERR: WRITE Pas de protocole d'emission pour " & Objet.Name)
            '            Exit Sub
            '        Case Else
            '            WriteLog("ERR: WRITE Protocole non géré : " & Objet.Modele.ToString.ToUpper)
            '            Exit Sub
            '    End Select
            'End If

            WriteLog("DBG: Commande passée à l arduino : " & MySensorsCommand)
            serialPortObj.WriteLine(MySensorsCommand) ', 0, 8)

        Catch ex As Exception
            WriteLog("ERR: WRITE " & ex.ToString)
        End Try
    End Sub

    ''' <summary>Fonction lancée lors de la suppression d'un device</summary>
    ''' <param name="DeviceId">Objet représetant le device à interroger</param>
    ''' <remarks></remarks>
    Public Sub DeleteDevice(ByVal DeviceId As String) Implements HoMIDom.HoMIDom.IDriver.DeleteDevice
        Try

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " DeleteDevice", ex.Message)
        End Try
    End Sub

    ''' <summary>Fonction lancée lors de l'ajout d'un device</summary>
    ''' <param name="DeviceId">Objet représetant le device à interroger</param>
    ''' <remarks></remarks>
    Public Sub NewDevice(ByVal DeviceId As String) Implements HoMIDom.HoMIDom.IDriver.NewDevice
        Try

        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " NewDevice", ex.Message)
        End Try
    End Sub

    ''' <summary>ajout des commandes avancées pour les devices</summary>
    ''' <remarks></remarks>
    Private Sub add_devicecommande(ByVal nom As String, ByVal description As String, ByVal nbparam As Integer)
        Try
            Dim x As New DeviceCommande
            x.NameCommand = nom
            x.DescriptionCommand = description
            x.CountParam = nbparam
            _DeviceCommandPlus.Add(x)
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " add_devicecommande", "Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>ajout Libellé pour le Driver</summary>
    ''' <param name="nom">Nom du champ : HELP</param>
    ''' <param name="labelchamp">Nom à afficher : Aide</param>
    ''' <param name="tooltip">Tooltip à afficher au dessus du champs dans l'admin</param>
    ''' <remarks></remarks>
    Private Sub Add_LibelleDriver(ByVal Nom As String, ByVal Labelchamp As String, ByVal Tooltip As String, Optional ByVal Parametre As String = "")
        Try
            Dim y0 As New HoMIDom.HoMIDom.Driver.cLabels
            y0.LabelChamp = Labelchamp
            y0.NomChamp = UCase(Nom)
            y0.Tooltip = Tooltip
            y0.Parametre = Parametre
            _LabelsDriver.Add(y0)
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " add_devicecommande", "Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Ajout Libellé pour les Devices</summary>
    ''' <param name="nom">Nom du champ : HELP</param>
    ''' <param name="labelchamp">Nom à afficher : Aide, si = "@" alors le champ ne sera pas affiché</param>
    ''' <param name="tooltip">Tooltip à afficher au dessus du champs dans l'admin</param>
    ''' <remarks></remarks>
    Private Sub Add_LibelleDevice(ByVal Nom As String, ByVal Labelchamp As String, ByVal Tooltip As String, Optional ByVal Parametre As String = "")
        Try
            Dim ld0 As New HoMIDom.HoMIDom.Driver.cLabels
            ld0.LabelChamp = Labelchamp
            ld0.NomChamp = UCase(Nom)
            ld0.Tooltip = Tooltip
            ld0.Parametre = Parametre
            _LabelsDevice.Add(ld0)
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " add_devicecommande", "Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>ajout de parametre avancés</summary>
    ''' <param name="nom">Nom du parametre (sans espace)</param>
    ''' <param name="description">Description du parametre</param>
    ''' <param name="valeur">Sa valeur</param>
    ''' <remarks></remarks>
    Private Sub add_paramavance(ByVal nom As String, ByVal description As String, ByVal valeur As Object)
        Try
            Dim x As New HoMIDom.HoMIDom.Driver.Parametre
            x.Nom = nom
            x.Description = description
            x.Valeur = valeur
            _Parametres.Add(x)
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " add_paramavance", "ERR: " & ex.Message)
        End Try
    End Sub

    ''' <summary>Creation d'un objet de type</summary>
    ''' <remarks></remarks>
    Public Sub New()
        Try
            _Version = Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString

            'Parametres avancés
            add_paramavance("Debug", "Activer le Debug complet (True/False)", False)
            add_paramavance("BaudRate", "Vitesse du port COM (57600 ou 9600)", 9600)
            add_paramavance("ErrorReceived", "Gérer les erreurs de réception (True=Activé, False=Désactivé)", True)
            'add_paramavance("AutoDiscover", "Permet de créer automatiquement des composants si ceux-ci n'existent pas encore (True/False)", False)

            'liste des devices compatibles
            _DeviceSupport.Add(ListeDevices.APPAREIL.ToString)
            _DeviceSupport.Add(ListeDevices.BAROMETRE.ToString)
            _DeviceSupport.Add(ListeDevices.BATTERIE.ToString)
            _DeviceSupport.Add(ListeDevices.COMPTEUR.ToString)
            _DeviceSupport.Add(ListeDevices.CONTACT.ToString)
            _DeviceSupport.Add(ListeDevices.DETECTEUR.ToString)
            _DeviceSupport.Add(ListeDevices.DIRECTIONVENT.ToString)
            _DeviceSupport.Add(ListeDevices.ENERGIEINSTANTANEE.ToString)
            _DeviceSupport.Add(ListeDevices.ENERGIETOTALE.ToString)
            _DeviceSupport.Add(ListeDevices.GENERIQUEBOOLEEN.ToString)
            _DeviceSupport.Add(ListeDevices.GENERIQUESTRING.ToString)
            _DeviceSupport.Add(ListeDevices.GENERIQUEVALUE.ToString)
            _DeviceSupport.Add(ListeDevices.HUMIDITE.ToString)
            _DeviceSupport.Add(ListeDevices.LAMPE.ToString)
            _DeviceSupport.Add(ListeDevices.PLUIECOURANT.ToString)
            _DeviceSupport.Add(ListeDevices.PLUIETOTAL.ToString)
            _DeviceSupport.Add(ListeDevices.SWITCH.ToString)
            _DeviceSupport.Add(ListeDevices.TELECOMMANDE.ToString)
            _DeviceSupport.Add(ListeDevices.TEMPERATURE.ToString)
            _DeviceSupport.Add(ListeDevices.TEMPERATURECONSIGNE.ToString)
            _DeviceSupport.Add(ListeDevices.UV.ToString)
            _DeviceSupport.Add(ListeDevices.VITESSEVENT.ToString)
            _DeviceSupport.Add(ListeDevices.VOLET.ToString)

            'ajout des commandes avancées pour les devices
            'add_devicecommande("COMMANDE", "DESCRIPTION", nbparametre)
            'add_devicecommande("CONFIG_TYPE_PIN", "configurer le type de PIN sur l arduino suivant les propriétés du composant", 0)
            'add_devicecommande("PWM", "Envoyer une commande PWM avec une valeur de 0 à 255", 1)
            'add_devicecommande("SETVAR", "Envoyer une valeur de type string à une variable sur l arduino", 1)
            'add_devicecommande("READX", "Lire les valeurs de toutes les entrées de l'arduino et mettre tous les composants Homidom à jour", 1)
            'add_devicecommande("SETLCD", "Ecrire un texte sur un ecran LCD.", 2)

            'Libellé Driver
            Add_LibelleDriver("HELP", "Aide...", "Pas d'aide actuellement...")

            'Libellé Device
            Add_LibelleDevice("ADRESSE1", "ID du noeud", "Valeur de type numérique")
            Add_LibelleDevice("ADRESSE2", "ID du capteur ou pin", "Valeur de type numérique")
            Add_LibelleDevice("SOLO", "@", "")
            Add_LibelleDevice("MODELE", "TYPE MySensors", "Détail des types dans la documentation du driver", "V_ARMED|V_CURRENT|V_DIMMER|V_DIRECTION|V_DISTANCE|V_DOWN|V_DUST_LEVEL|V_FLOWVOLUME|V_FORECAST|V_GUST|V_HEATER|V_HEATER_SW|V_HUM|V_IMPEDANCE|V_IR_SEND|V_IR_RECEIVE|V_LIGHT|V_LIGHT_LEVEL|V_LOCK_STATUS|V_KWH|V_PRESSURE|V_RAIN|V_RAINRATE|V_SCENE_ON|V_SCENE_OFF|V_TEMP|V_STOP|V_TRIPPED|V_UP|V_UV|V_VAR1|V_VAR2|V_VAR3|V_VAR4|V_VAR5|V_VOLTAGE|V_WATT|V_WEIGHT|V_WIND")
            Add_LibelleDevice("REFRESH", "Refresh", "0")
            'Add_LibelleDevice("LASTCHANGEDUREE", "LastChange Durée", "")
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " New", ex.Message)
        End Try
    End Sub

    ''' <summary>Si refresh >0 gestion du timer</summary>
    ''' <remarks>PAS UTILISE CAR IL FAUT LANCER UN TIMER QUI LANCE/ARRETE CETTE FONCTION dans Start/Stop</remarks>
    Private Sub TimerTick(ByVal source As Object, ByVal e As System.Timers.ElapsedEventArgs)

    End Sub

#End Region

#Region "Fonctions internes"

    Private Sub serialPortObj_ErrorReceived(ByVal sender As Object, ByVal e As SerialErrorReceivedEventArgs)
        Select Case e.EventType
            Case SerialError.Frame
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " ErrorReceived", "Error: Le matériel a détecté une erreur de trame")
            Case SerialError.Overrun
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " ErrorReceived", "Error: Un dépassement de mémoire tampon de caractères s'est produit.Le caractère suivant est perdu")
            Case SerialError.RXOver
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " ErrorReceived", "Error: Un dépassement de la mémoire tampon d'entrée s'est produit.Il n'y a plus de place dans la mémoire tampon d'entrée ou un caractère a été reçu après le caractère de fin de fichier")
                serialPortObj.DiscardInBuffer()
            Case SerialError.RXParity
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " ErrorReceived", "Error: Le matériel a détecté une erreur de parité")
            Case SerialError.TXFull
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " ErrorReceived", "Error: L'application a essayé de transmettre un caractère, mais la mémoire tampon de sortie était pleine")
            Case Else
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " ErrorReceived", "Erreur inconnue, le driver va tenter de traiter les données")
                Dim line As String = serialPortObj.ReadLine()
                _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " ErrorReceived", "Données reçues: " & line)

        End Select
    End Sub

    ''' <summary>
    ''' Traite les infos reçus
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub DataReceived(ByVal sender As Object, ByVal e As SerialDataReceivedEventArgs)
        Try
            'If first Then
            'first = False
            'Else
            'on attend d'avoir le reste
            'System.Threading.Thread.Sleep(500)

            '                Dim line As String = serialPortObj.ReadExisting
            serialPortObj.ReadTimeout = 1000
            Do
                Dim line As String = serialPortObj.ReadLine()
                If line Is Nothing Then
                    Exit Do
                Else
                    line = line.Replace(vbCr, "").Replace(vbLf, "")
                    _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " DataReceived", "Données reçues: " & line)
                    Dim aryLine() As String
                    aryLine = line.Split(";")
                    ' Action après réception d'une trame sur le port COM/USB
                    If UBound(aryLine) >= 5 Then
                        Dim Commande As String = aryLine(3)
                        Dim Valeur As String = aryLine(5)
                        'If aryLine(0) = "DEBUG" Then
                        '    _Server.Log(TypeLog.INFO, TypeSource.DRIVER, Me.Nom & " Datareceived", "From Arduino : " & line)
                        'Else
                        Select Case aryLine(2)
                            Case "0" ' Message type "presentation"
                                _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Message Type 'presentation' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                Select Case aryLine(4)
                                    Case "0"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_DOOR' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "1"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_MOTION' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "2"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_SMOKE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "3"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_LIGHT' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "4"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_DIMMER' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "5"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_COVER' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "6"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_TEMP' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "7"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_HUM' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "8"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_BARO' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "9"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_WIND' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "10"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_RAIN' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "11"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_UV' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "12"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_WEIGHT' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "13"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_POWER' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "14"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_HEATE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "15"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_DISTANCE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "16"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_LIGHT_LEVEL' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "17"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_ARDUINO_NODE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "18"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_ARDUINO_RELAY' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "19"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_LOCK' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "20"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_IR' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "21"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_WATER' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "22"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_AIR_QUALITY' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "23"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_CUSTOM' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "24"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_DUST' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "25"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_SCENE_CONTROLLER' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case Else
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'S_????' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                End Select
                            Case "1", "2" ' Message type "set/req"
                                _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Message Type 'set/req' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                Select Case aryLine(4)
                                    Case "0"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_TEMP' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "1"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_HUM' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "2"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_LIGHT' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "3"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_DIMMER' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "4"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_PRESSURE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "5"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_FORECAST' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "6"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_RAIN' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "7"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_RAINRATE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "8"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_WIND' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "9"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_GUST' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "10"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_DIRECTION' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "11"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_UV' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "12"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_WEIGHT' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "13"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_DISTANCE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "14"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_IMPEDANCE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "15"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_ARMED' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "16"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_TRIPPED' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "17"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_WATT' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "18"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_KWH' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "19"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_SCENE_ON' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "20"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_SCENE_OFF' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "21"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_HEATER' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "22"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_HEATER_SW' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "23"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_LIGHT_LEVEL' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "24"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_VAR1' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "25"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_VAR2' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "26"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_VAR3' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "27"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_VAR4' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "28"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_VAR5' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "29"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_UP' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "30"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_DOWN' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "31"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_STOP' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "32"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_IR_SEND' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "33"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_IR_RECEIVE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "34"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_FLOW' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "35"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_VOLUME' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "36"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_LOCK_STATUS' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "37"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_DUST_LEVEL' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "38"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_VOLTAGE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case "39"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_CURRENT' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                        traitement(aryLine(2), aryLine(4), aryLine(0), aryLine(1), aryLine(5))
                                    Case Else
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'V_????' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                End Select

                            Case "3" ' Message type "internal"
                                _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Message Type 'internal' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                Select Case aryLine(4)
                                    Case "0"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_BATTERY_LEVEL' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "1"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_TIME' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "2"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_VERSION' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "3"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_ID_REQUEST' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "4"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_ID_RESPONSE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "5"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_INCLUSION_MODE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "6"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_CONFIG' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "7"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_FIND_PARENT' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "8"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_FIND_PARENT_RESPONSE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "9"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_LOG_MESSAGE' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "10"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_CHILDREN' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "11"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_SKETCH_NAME' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "12"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_SKETCH_VERSION' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "13"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_REBOOT' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case "14"
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_GATEWAY_READY' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                    Case Else
                                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Type 'I_????' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                                End Select
                            Case "4" ' Message type "stream"
                                _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Message Type 'stream' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))
                            Case Else
                                _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Datareceived", "Message Type '????' " & aryLine(0) & ";" & aryLine(1) & ";" & aryLine(2) & ";" & aryLine(3) & ";" & aryLine(4) & ";" & aryLine(5))

                        End Select
                        'End If
                    End If
                    'End If
                End If
            Loop

        Catch Ex As Exception
            '_Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Datareceived", "Erreur:" & Ex.ToString)
        End Try
    End Sub

    ''' <summary>Traite les paquets reçus</summary>
    ''' <remarks></remarks>
    Private Sub traitement(ByVal msgtype As String, ByVal type As String, ByVal adresse As String, ByVal adresse2 As String, ByVal valeur As String)
        '    Private Sub traitement(ByVal adresse As String, ByVal valeur As String)
        Try
            'correction valeur
            valeur = Regex.Replace(valeur, "[.,]", System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)

            'Recherche si un device affecté
            Dim listedevices As New ArrayList
            Dim homidom_type As Integer
            Dim _Type As String = ""
            Dim autodevice As Boolean = True
            Dim deviceupdate As Boolean = False

            Select Case msgtype
                Case 0
                    '_Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Traitement : ", "Noeud " & adresse & " Sensor " & adresse2 & " Type " & msgtype & " Valeur " & valeur)
                    valeur = vbNull
                    Select Case type
                        Case 0 'S_DOOR
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 12
                        Case 1 'S_MOTION
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 12
                        Case 2 'S_SMOKE
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 12
                        Case 3 'S_LIGHT
                            _Type = "LAMPE"
                            homidom_type = 16
                        Case 4 'S_DIMMER
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 14
                        Case 5 'S_COVER
                            _Type = "VOLET"
                            homidom_type = 27
                        Case 6 'S_TEMP
                            _Type = "TEMPERATURE"
                            homidom_type = 22
                        Case 7 'S_HUM
                            _Type = "HUMIDITE"
                            homidom_type = 14
                        Case 8 'S_BARO
                            _Type = "BAROMETRE"
                            homidom_type = 3
                        Case 9 'S_WIND
                            _Type = "VITESSEVENT"
                            homidom_type = 26
                        Case 10 'S_RAIN
                            _Type = "PLUIECOURANT"
                            homidom_type = 19
                        Case 11 'S_UV
                            _Type = "UV"
                            homidom_type = 25
                        Case 12 'S_WEIGHT
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 14
                        Case 13 'S_POWER
                            _Type = "ENERGIEINSTANTANEE"
                            homidom_type = 9
                        Case 14 'S_HEATER
                            _Type = "GENERIQUESTRING"
                            homidom_type = 13
                        Case 15 'S_DISTANCE
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 14
                        Case 16 'S_LIGHT_LEVEL
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 14
                        Case 17 'S_ARDUINO_NODE
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 14
                            autodevice = False
                        Case 18 'S_ARDUINO_RELAY
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 14
                            autodevice = False
                        Case 19 'S_LOCK
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 12
                        Case 20 'S_IR
                            _Type = "GENERIQUESTRING"
                            homidom_type = 13
                        Case 21 'S_WATER
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 14
                        Case 22 'S_AIR_QUALITY
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 14
                        Case 23 'S_CUSTOM
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 14
                        Case 24 'S_DUST
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 14
                        Case 25 'S_SCENE_CONTROLLER
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 12
                    End Select
                Case 1, 2
                    '_Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Traitement : ", "Noeud " & adresse & " Sensor " & adresse2 & " Type " & msgtype & " Valeur " & valeur)
                    deviceupdate = True
                    Select Case type
                        Case 0 'V_TEMP (température)
                            _Type = "TEMPERATURE"
                            homidom_type = 22
                        Case 1 'V_HUM (pourcentage d'humidité)
                            _Type = "HUMIDITE"
                            homidom_type = 14
                        Case 2 'V_LIGHT (etat de la lumière on-off : 0=Off, 1=On)
                            _Type = "LAMPE"
                            homidom_type = 15
                        Case 3 'V_DIM (valeur du variateur 0-100%)
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 4 'V_PRESSURE (pression atmosphérique)
                            _Type = "BAROMETRE"
                            homidom_type = 2
                        Case 5 'V_FORECAST (Prévisions météo stable "stable", ensoleillé "sunny", orage ""thunderstorm", instable "unstable", nuageux "unstable" ou inconnu "unknown")
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 6 'V_RAIN (quantite de pluie)
                            _Type = "PLUIECOURANT"
                            homidom_type = 18
                        Case 7 'V_RAINRATE (intensité de la pluie)
                            _Type = "PLUIETOTAL"
                            homidom_type = 19
                        Case 8 'V_WIND (vitesse du vent)
                            _Type = "VITESSEVENT"
                            homidom_type = 25
                        Case 9 'V_GUST (type de vent)
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 10 'V_DIRECTION (direction du vent)
                            _Type = "DIRECTIONVENT"
                            homidom_type = 7
                        Case 11 'V_UV (indice d'UV)
                            _Type = "UV"
                            homidom_type = 24
                        Case 12 'V_WEIGHT (poids)
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 13
                        Case 13 'V_DISTANCE (mesure de distance)
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 13
                        Case 14 'V_IMPEDANCE (valeur de l'impedance)
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 13
                        Case 15 'V_ARMED (état armé d'un capteur de sécurité armé/ignoré : 1=Armed, 0=Bypassed)
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 11
                        Case 16 'V_TRIPPED (état déclenché par un capteur de sécurité déclenché/non déclenché : 1=Tripped, 0=Untripped)
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 11
                        Case 17 'V_WATT (valeur en watts pour les compteurs électriques)
                            _Type = "ENERGIEINSTANTANEE"
                            homidom_type = 8
                        Case 18 'V_KMH (nombre cumulé de KW/h)
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 13
                        Case 19 'V_SCENE_ON (activer un sénario)
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 11
                        Case 20 'V_SCENE_OFF (désactiver un sénario)
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 11
                        Case 21 'V_HEATER (mode de chauffage : arrêt "Off", chaud "HeatOn", froid "CoolOn", changement automatique)
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 22 'V_HEATER_SW (interrupteur d'alimentation du chauffage : 1=On, 0=Off)
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 11
                        Case 23 'V_LIGHT_LEVEL (niveau de lumière 0-100%)
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 13
                        Case 24 'V_VAR1 (valeur personnalisée N°1)
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 25 'V_VAR2 (valeur personnalisée N°2)
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 26 'V_VAR3 (valeur personnalisée N°3)
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 27 'V_VAR4 (valeur personnalisée N°4)
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 28 'V_VAR5 (valeur personnalisée N°5)
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 29 'V_UP (commande de volet "Up")
                            _Type = "VOLET"
                            homidom_type = 26
                        Case 30 'V_DOWN (commande de volet "Down")
                            _Type = "VOLET"
                            homidom_type = 26
                        Case 31 'V_STOP (commande de volet "Stop")
                            _Type = "VOLET"
                            homidom_type = 26
                        Case 32 'V_IR_SEND (envoi d'une commande Infrarouge)
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 33 'V_IR_RECEIVE (reception d'une commande Infrarouge)
                            _Type = "GENERIQUESTRING"
                            homidom_type = 12
                        Case 34 'V_FLOW (niveau d'eau en mètre)
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 13
                        Case 35 'V_VOLUME (volume d'eau)
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 13
                        Case 36 'V_LOCK_STATUS (status de vérouillage : 1=Locked, 0=Unlocked)
                            _Type = "GENERIQUEBOOLEEN"
                            homidom_type = 11
                        Case 37 'V_DUST_LEVEL ()
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 13
                        Case 38 'V_VOLTAGE (tension mesurée)
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 13
                        Case 39 'V_CURRENT (courant mesurée)
                            _Type = "GENERIQUEVALUE"
                            homidom_type = 13
                        Case Else
                            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Process", "Le type de device n'appartient pas à ce driver: " & type)
                            Exit Sub
                    End Select
            End Select

            listedevices = _Server.ReturnDeviceByAdresse1TypeDriver(_idsrv, adresse, _Type, Me._ID, True)
            'un device trouvé on maj la valeur
            If (listedevices.Count = 1) Then
                If deviceupdate = True Then
                    listedevices.Item(0).Value = valeur
                    _Server.Log(TypeLog.INFO, TypeSource.DRIVER, Me.Nom & " Reception : ", "Noeud N° " & adresse & " Capteur/Actionneur " & adresse2 & " Valeur " & valeur)
                End If
            ElseIf (listedevices.Count > 1) Then
                For i As Integer = 0 To listedevices.Count - 1
                    If listedevices.Item(i).adresse2.ToUpper() = adresse2.ToUpper() Then
                        If deviceupdate = True Then
                            listedevices.Item(i).Value = valeur
                            _Server.Log(TypeLog.INFO, TypeSource.DRIVER, Me.Nom & " Reception : ", "Noeud N° " & adresse & " Capteur/Actionneur " & adresse2 & " Valeur " & valeur)
                        End If
                    End If
                Next
                '_Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Process", "Plusieurs devices correspondent à : " & adresse)
            Else
                'si autodiscover = true ou modedecouverte du serveur actif alors on crée le composant sinon on logue
                If autodevice = True Then
                    If (_AutoDiscover Or _Server.GetModeDecouverte) Then
                        _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom & " Process", "Device non trouvé, AutoCreation du composant : " & _Type & " " & adresse & " " & adresse2 & ":" & valeur)
                        _Server.AddDetectNewDevice(adresse, _ID, homidom_type, adresse2, valeur)
                    Else
                        _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " Process", "Device non trouvé : " & _Type & " " & adresse & " " & adresse2 & ":" & valeur)
                    End If
                End If
            End If
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " traitement", "Exception : " & ex.Message & " --> " & adresse & " " & adresse2 & " : " & valeur)
        End Try
    End Sub

    Private Sub WriteLog(ByVal message As String)
        Try
            'utilise la fonction de base pour loguer un event
            If STRGS.InStr(message, "DBG:") > 0 Then
                _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom, STRGS.Right(message, message.Length - 5))
            ElseIf STRGS.InStr(message, "ERR:") > 0 Then
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom, STRGS.Right(message, message.Length - 5))
            Else
                _Server.Log(TypeLog.INFO, TypeSource.DRIVER, Me.Nom, message)
            End If
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " WriteLog", ex.Message)
        End Try
    End Sub

#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

