Attribute VB_Name = "Players"
Global Personagem(1000) As PersonagemInfo

Public Type PersonagemInfo
IP As String
Porta As String

Login As String
Char(4) As String

Nome As String
Level As Integer
XP As Integer
Classe As Integer
Zen As Integer
Mapa As Integer
PosX As Integer
PosY As Integer
Inventario As Byte
Magias As Byte
Quest As Byte

End Type

Public Function ZerarPlayerInfo(Index As Integer)
Personagem(Index).IP = ""
Personagem(Index).Porta = ""

Personagem(Index).Login = ""
Personagem(Index).Char(0) = ""
Personagem(Index).Char(1) = ""
Personagem(Index).Char(2) = ""
Personagem(Index).Char(3) = ""
Personagem(Index).Char(4) = ""

Personagem(Index).Nome = ""
Personagem(Index).Level = ""
Personagem(Index).XP = ""
Personagem(Index).Classe = ""
Personagem(Index).Zen = ""
Personagem(Index).Mapa = ""
Personagem(Index).PosX = ""
Personagem(Index).PosY = ""
Personagem(Index).Inventario = &HF
Personagem(Index).Magias = &HF
Personagem(Index).Quest = &HF
End Function
