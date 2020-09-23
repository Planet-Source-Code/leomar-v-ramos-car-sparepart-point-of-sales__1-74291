Attribute VB_Name = "ModVarTypes"
Public Type USER_INFO
    USERID                              As String
    USERNAME                            As String
    PASSWORD                            As String
    FULLNAME                            As String
    USERTYPE                            As String
    USER_ISADMIN                        As Boolean
End Type

Public Enum FORM_STATE
    AddStateMode = 0
    EditStateMode = 1
End Enum

Public Type COMPANY_INFO
    COMPANYID                           As String
    COMPANYNAME                         As String
    ADDRESS                             As String
    BUSINESSNO                          As String
    FAXNO                               As String
    EMAIL                               As String
End Type
