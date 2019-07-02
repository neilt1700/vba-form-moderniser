Attribute VB_Name = "MainModule"
' Copyright (c) Commtap CIC 2019
' Available under the MIT license: see the LICENSE file at the root of this
' project.
' Contact: tap@commtap.org

Option Explicit

Public Sub ShowSampleForm()

  Dim oUSampleUserForm As USampleUserForm
  Set oUSampleUserForm = Factory.CreateUSampleUserForm
  
  ' Modernising
  Set FormModerniserModule.gb_colCurrentUserForms = New Collection
  FormModerniserModule.gb_colCurrentUserForms.Add oUSampleUserForm
  FormModerniserModule.ModerniseForm oUSampleUserForm
  
  oUSampleUserForm.Show

End Sub
