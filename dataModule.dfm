object DataModule1: TDataModule1
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 396
  Width = 401
  object ConexionBD: TFDConnection
    Params.Strings = (
      'DriverID=MSAcc'
      
        'Database=D:\Proyectos Field-Map\Testing\Proyecto_Diagnostico\Fie' +
        'ldMapData_Proyecto_Diagnostico.accdb')
    Connected = True
    Left = 192
    Top = 24
  end
  object SQLQuery: TFDQuery
    Connection = ConexionBD
    Left = 176
    Top = 296
  end
  object FDPhysFBDriverLink1: TFDPhysFBDriverLink
    Left = 64
    Top = 60
  end
  object FDPhysMSAccessDriverLink1: TFDPhysMSAccessDriverLink
    Left = 304
    Top = 64
  end
end
