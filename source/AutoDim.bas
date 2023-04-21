Attribute VB_Name = "AutoDim"
'===============================================================================
'
'    VBA MACRO For CORELDRAW / COTATIONS AUTOMATIQUES
'    Copyright (C) 2020 Fabrice VAN NEER
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.
'    Tradução básica p/ português @corelnaveia 2021/04/29
'
'    Atualização 2023 Ferreira Felipe
'    Atualização 04/2023 compatibilidade com X6 + Idioma RU por Elvin Macros!
'
'===============================================================================

Option Explicit

Public Const APP_NAME As String = "AutoDimensions"
Public Const APP_VERSION = "2023"

Public Const DIMENSIONS_COLOR As String = "CMYK,USER,100,20,0,0"
Public Const DIMENSIONS_FONT As String = "ARIAL"
Public Const DIMENSIONS_STR As String = "Dimension"

'===============================================================================

Sub Start()
    With New MainView
        .Show vbModeless
    End With
End Sub
