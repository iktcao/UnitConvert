Public Class Converter
    Private strATM As String

    Public Sub New()
        strATM = "101.325"
    End Sub

    Public Property ATM() As String
        Get
            Return strATM
        End Get
        Set(value As String)
            If value <> "" Then
                If CType(value, Double) < 0 Then
                    MessageBox.Show("输入的绝对压力值不能小于0，请重新输入！"， "错误提示", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    strATM = value
                End If
            End If
        End Set
    End Property

    ''' <summary>
    ''' 转换为该单位类型下的所有单位，返回一个字典
    ''' </summary>
    ''' <param name="unit">原单位名称</param>
    ''' <param name="value">原单位数值</param>
    ''' <param name="unitset">单位类型</param>
    ''' <returns></returns>
    Public Function Convert(ByVal unit As String, ByVal value As String, ByVal unitset As String) As Dictionary(Of String, String)

        Dim dblValueInput As Double
        Dim dblValueSI As Double
        Dim dblFactor As Double
        Dim dblSeed As Double

        Dim dicUnitSet As New Dictionary(Of String, String())
        Dim dicConvertUnits As New Dictionary(Of String, String)

        dblValueInput = CType(value, Double)

        dicUnitSet = GetDict(CType([Enum].Parse(GetType(UnitSet), unitset), UnitSet))

        If unitset = "Pressure" Then
            dblFactor = EvalATM(dicUnitSet(unit)(1))
            dblSeed = EvalATM(dicUnitSet(unit)(2))
        Else
            dblFactor = CType(dicUnitSet(unit)(1), Double)
            dblSeed = CType(dicUnitSet(unit)(2), Double)
        End If

        dblValueSI = (dblValueInput + dblSeed) * dblFactor

        For Each key As String In dicUnitSet.Keys
            If unitset = "Pressure" Then
                dblFactor = EvalATM(dicUnitSet(key)(1))
                dblSeed = EvalATM(dicUnitSet(key)(2))
            Else
                dblFactor = CType(dicUnitSet(key)(1), Double)
                dblSeed = CType(dicUnitSet(key)(2), Double)
            End If
            Dim item As Double = dblValueSI / dblFactor - dblSeed
            dicConvertUnits.Add(key, item.ToString())
        Next

        Return dicConvertUnits

    End Function

    ''' <summary>
    ''' 转换为指定单位，返回指定单位数值的文本
    ''' </summary>
    ''' <param name="unit">原单位名称</param>
    ''' <param name="value">原单位数值</param>
    ''' <param name="unitset">单位类型</param>
    ''' <param name="targetunit">目标单位名称</param>
    ''' <returns></returns>
    Public Function Convert(ByVal unit As String, ByVal value As String, ByVal unitset As String, ByVal targetunit As String) As String
        Dim dicConvertUnits As New Dictionary(Of String, String)
        dicConvertUnits = Convert(unit, value, unitset)
        Return dicConvertUnits(targetunit)
    End Function

    Public Function EvalATM(ByVal str As String) As Double
        If str.Contains("atm") Then
            If str = "atm" Then
                Return CType(strATM, Double)
            Else
                If str.Substring(3, 1) = "*" Then
                    Return CType(strATM, Double) * CType(str.Replace("atm*", ""), Double)
                Else
                    Return CType(strATM, Double) / CType(str.Replace("atm/", ""), Double)
                End If
            End If
        Else
            Return CType(str, Double)
        End If
    End Function


    Private Function GetTemperature() As Dictionary(Of String, String())
        Dim dicTemperature As New Dictionary(Of String, String())
        dicTemperature.Add("K", {"Kelvin", "1", "0.0"})
        dicTemperature.Add("F", {"Fahrenheit", "0.555555555555555", "459.67"})
        dicTemperature.Add("C", {"Celsius", "1", "273.15"})
        dicTemperature.Add("R", {"Rankine", "0.555555555555555", "0.0"})
        Return dicTemperature
    End Function

    Private Function GetPressure() As Dictionary(Of String, String())
        Dim dicPressure As New Dictionary(Of String, String())
        dicPressure.Add("Pa", {"pascal absolute", "1.0", "0.0"})
        dicPressure.Add("PaG", {"pascal gauge", "1.0", "atm*1000"})
        dicPressure.Add("kPa", {"kilopascal absolute", "1000", "0.0"})
        dicPressure.Add("kPaG", {"kilopascal gauge", "1000", "atm"})
        dicPressure.Add("bar", {"bar absolute", "100000", "0.0"})
        dicPressure.Add("barG", {"bar gauge", "100000", "atm/100"})
        dicPressure.Add("kgf__cm2", {"Kgf/cm2 absolute", "98066.52", "0.0"})
        dicPressure.Add("kgf__cm2G", {"Kgf/cm2 gauge", "98066.52", "atm/98.06652"})
        dicPressure.Add("MPa", {"megapascal absolute", "1000000", "0.0"})
        dicPressure.Add("MPaG", {"megapascal gauge", "1000000", "atm/1000"})
        dicPressure.Add("atm", {"standard atmosphere abosulte", "atm*1000", "0.0"})
        dicPressure.Add("atmG", {"standard atmosphere gauge", "atm*1000", "atm"})
        dicPressure.Add("mbar", {"millibar absolute", "0.001", "0.0"})
        dicPressure.Add("mbarG", {"millibar gauge", "0.001", "atm/0.000001"})
        dicPressure.Add("mmHg", {"millimeters of Hg (0.0C) absolute", "133.322", "0.0"})
        dicPressure.Add("mmHgG", {"millimeters of Hg (0.0C) gauge", "133.322", "atm/0.133322"})
        dicPressure.Add("inHg", {"inches of Hg (0.0C) absolute", "3386.387", "0.0"})
        dicPressure.Add("inHgG", {"inches of Hg (0.0C) gauge", "3386.387", "atm/3.386387"})
        dicPressure.Add("inH2O", {"inches of H2O (4.0C) absolute", "249.082", "0.0"})
        dicPressure.Add("inH2OG", {"inches of H2O (4.0C) gauge", "249.082", "atm/0.249082"})
        dicPressure.Add("mH2O", {"meters of H2O (4.0C) absolute", "9806.652", "0.0"})
        dicPressure.Add("mH2OG", {"meters of H2O (4.0C) gauge", "9806.652", "atm/9.806652"})
        dicPressure.Add("Torr", {"torr", "133.322", "0.0"})
        dicPressure.Add("psi", {"pounds/square inch absolute", "6894.757", "0.0"})
        dicPressure.Add("psiG", {"pounds/square inch gauge", "6894.757", "atm/6.894757"})
        dicPressure.Add("psf", {"pounds/square foot absolute", "47.88025898", "0.0"})
        dicPressure.Add("psfG", {"pounds/square foot gauge", "47.88025898", "atm/0.04788025898"})
        Return dicPressure
    End Function

    Private Function GetMass() As Dictionary(Of String, String())
        Dim dicMass As New Dictionary(Of String, String())
        dicMass.Add("kg", {"kilogram", "1.0", "0.0"})
        dicMass.Add("g", {"gram", "0.00100", "0.0"})
        dicMass.Add("mg", {"milligram", "0.000001", "0.0"})
        dicMass.Add("ug", {"microgram", "0.000000001", "0.0"})
        dicMass.Add("ton", {"metric ton or tonne", "1000.0", "0.0"})
        dicMass.Add("kton", {"kilotons", "1.0E+6", "0.0"})
        dicMass.Add("Mton", {"megatons", "1.0E+9", "0.0"})
        dicMass.Add("MMton", {"gigatons", "1.0E+12", "0.0"})
        dicMass.Add("lb", {"pound", "0.45359237", "0.0"})
        dicMass.Add("klb", {"kilopounds", "453.59237", "0.0"})
        dicMass.Add("ston", {"short ton or US tons", "907.18", "0.0"})
        dicMass.Add("lton", {"long ton or UK tons", "1016.0", "0.0"})
        dicMass.Add("oz", {"ounce", "0.028350", "0.0"})
        dicMass.Add("ct", {"carat", "0.0002", "0.0"})
        Return dicMass
    End Function


    Private Function GetLength() As Dictionary(Of String, String())
        Dim dicLength As New Dictionary(Of String, String())
        dicLength.Add("m", {"meter", "1.0", "0.0"})
        dicLength.Add("km", {"kilometer", "1000.0", "0.0"})
        dicLength.Add("dm", {"decimeter", "0.1", "0.0"})
        dicLength.Add("cm", {"centimeter", "0.01", "0.0"})
        dicLength.Add("mm", {"millimeter", "0.001", "0.0"})
        dicLength.Add("um", {"micrometer", "1.0E-6", "0.0"})
        dicLength.Add("nm", {"nanometer", "1.0E-9", "0.0"})
        dicLength.Add("ft", {"feet", "0.3048", "0.0"})
        dicLength.Add("in", {"inch", "0.0254", "0.0"})
        dicLength.Add("yd", {"yard", "0.91440", "0.0"})
        dicLength.Add("mi", {"UK mile", "1.6093e+3", "0.0"})
        dicLength.Add("nmi", {"nautical mile", "1852.0", "0.0"})
        Return dicLength
    End Function

    Private Function GetTime() As Dictionary(Of String, String())
        Dim dicTime As New Dictionary(Of String, String())
        dicTime.Add("s", {"second", "1.0", "0.0"})
        dicTime.Add("ms", {"millisecond", "1.0E-3", "0.0"})
        dicTime.Add("min", {"minute", "60", "0.0"})
        dicTime.Add("hr", {"hour", "3600", "0.0"})
        dicTime.Add("day", {"day", "86400", "0.0"})
        dicTime.Add("wk", {"week", "6.0480E+5", "0.0"})
        dicTime.Add("mon", {"month", "2.5514E+6", "0.0"})
        dicTime.Add("yr", {"year", "3.1557E+7", "0.0"})
        Return dicTime
    End Function

    Private Function GetAngle() As Dictionary(Of String, String())
        Dim dicAngle As New Dictionary(Of String, String())
        dicAngle.Add("arcsec", {"second of arc", "1.0", "0.0"})
        dicAngle.Add("arcmin", {"minute of arc", "60.0", "0.0"})
        dicAngle.Add("deg", {"degree", "3600.0", "0.0"})
        dicAngle.Add("rad", {"radian", "206264.808", "0.0"})
        Return dicAngle
    End Function

    Private Function GetArea() As Dictionary(Of String, String())
        Dim dicArea As New Dictionary(Of String, String())
        dicArea.Add("m2", {"square meter", "1.0", "0.0"})
        dicArea.Add("km2", {"square kilometer", "1.0e+6", "0.0"})
        dicArea.Add("dm2", {"square decimeter", "1.0e-2", "0.0"})
        dicArea.Add("cm2", {"square centimeter", "1.0e-4", "0.0"})
        dicArea.Add("mm2", {"square millimeter", "1.0e-6", "0.0"})
        dicArea.Add("ft2", {"square feet", "9.2903e-2", "0.0"})
        dicArea.Add("in2", {"square inch", "6.4516e-4", "0.0"})
        dicArea.Add("yd2", {"square yard", "8.3613e-1", "0.0"})
        dicArea.Add("mi2", {"square mile", "2.59e+6", "0.0"})
        Return dicArea
    End Function

    Private Function GetVolume() As Dictionary(Of String, String())
        Dim dicVolume As New Dictionary(Of String, String())
        dicVolume.Add("m3", {"cubic meter", "1.0", "0.0"})
        dicVolume.Add("L", {"liter", "1.0e-3", "0.0"})
        dicVolume.Add("dm3", {"cubic decimeter", "1.0e-3", "0.0"})
        dicVolume.Add("mL", {"milliliter", "1.0e-6", "0.0"})
        dicVolume.Add("cm3", {"cubic centimeter", "1.0e-6", "0.0"})
        dicVolume.Add("bbl", {"barrel", "1.5899e-1", "0.0"})
        dicVolume.Add("ft3", {"cubic foot", "2.8317e-2", "0.0"})
        dicVolume.Add("in3", {"cubic inch", "1.6387e-5", "0.0"})
        dicVolume.Add("gal", {"US gallon", "3.7854e-3", "0.0"})
        dicVolume.Add("pt", {"US pint", "4.7318e-4", "0.0"})
        dicVolume.Add("oz", {"US ounce", "2.9574e-5", "0.0"})
        dicVolume.Add("igal", {"UK gallon", "4.5461e-3", "0.0"})
        dicVolume.Add("ipt", {"UK pint", "5.6826e-4", "0.0"})
        dicVolume.Add("ioz", {"UK ounce", "2.8413e-5", "0.0"})
        Return dicVolume
    End Function

    Private Function GetDensity() As Dictionary(Of String, String())
        Dim dicDensity As New Dictionary(Of String, String())
        dicDensity.Add("kg__m3", {"kilogram/cubic meter", "1.0", "0.0"})
        dicDensity.Add("g__cm3", {"gram/cubic centimeter", "1.0E+3", "0.0"})
        dicDensity.Add("sp", {"specific gravity", "999.972", "0.0"})
        dicDensity.Add("lb__ft3", {"pounds/cubic foot", "16.01845115", "0.0"})
        dicDensity.Add("lb__gal", {"pounds/gallon", "119.826", "0.0"})
        dicDensity.Add("lb__in3", {"pounds/cubic inch", "27679.883", "0.0"})
        dicDensity.Add("lb__barrel", {"pounds/barrel", "2.853", "0.0"})
        Return dicDensity
    End Function

    Private Function GetForce() As Dictionary(Of String, String())
        Dim dicForce As New Dictionary(Of String, String())
        dicForce.Add("N", {"newton", "1.0", "0.0"})
        dicForce.Add("kN", {"kilonewton", "1000.0", "0.0"})
        dicForce.Add("gf", {"gram force", "9.80665E-3", "0.0"})
        dicForce.Add("kgf", {"kilogram force", "9.80665", "0.0"})
        dicForce.Add("lbf", {"pound force", "4.448222", "0.0"})
        dicForce.Add("klbf", {"kilopounds force", "4.448222E+3", "0.0"})
        dicForce.Add("dyn", {"kilocalorie per second", "1.0E-5", "0.0"})
        Return dicForce
    End Function

    Private Function GetEnergy() As Dictionary(Of String, String())
        Dim dicEnergy As New Dictionary(Of String, String())
        dicEnergy.Add("J", {"joule", "1.0", "0.0"})
        dicEnergy.Add("kJ", {"kilojoule", "1000.0", "0.0"})
        dicEnergy.Add("MJ", {"megajoule", "1.0E+6", "0.0"})
        dicEnergy.Add("cal", {"calorie", "4.184", "0.0"})
        dicEnergy.Add("kcal", {"kilocalorie", "4184", "0.0"})
        dicEnergy.Add("Mcal", {"megacalorie", "4.184E+6", "0.0"})
        dicEnergy.Add("Btu", {"British thermal unit", "1055.0", "0.0"})
        dicEnergy.Add("hph", {"horsepower hour", "2684519.5392", "0.0"})
        dicEnergy.Add("kWh", {"kilowatt hour", "3600000.0", "0.0"})
        dicEnergy.Add("kgm", {"kilogram-force meter", "9.80665", "0.0"})
        dicEnergy.Add("gcm", {"gram-force centimeter", "0.000098067", "0.0"})
        Return dicEnergy
    End Function

    Private Function GetWork() As Dictionary(Of String, String())
        Dim dicWork As New Dictionary(Of String, String())
        dicWork.Add("W", {"watt", "1.0", "0.0"})
        dicWork.Add("kW", {"kilowatt", "1000.0", "0.0"})
        dicWork.Add("MW", {"megaWatts", "1000000.0", "0.0"})
        dicWork.Add("kcal__s", {"kilocalorie per second", "4184", "0.0"})
        dicWork.Add("kcal__h", {"kilocalorie/hour", "1.1627782", "0.0"})
        dicWork.Add("Btu__s", {"btu per second", "1055.0", "0.0"})
        dicWork.Add("Btu__h", {"btu per hour", "0.2930556", "0.0"})
        dicWork.Add("hp", {"horse power (U.K.)", "735.499", "0.0"})
        dicWork.Add("kgm__s", {"kilogram meter per second", "9.80665", "0.0"})
        dicWork.Add("Nm__s", {"newton meter per second", "1.0", "0.0"})
        dicWork.Add("ftlb__s", {"feet pound per second", "1.3557484", "0.0"})
        Return dicWork
    End Function

    Private Function GetThermalConductivity() As Dictionary(Of String, String())
        Dim dicThermalConductivity As New Dictionary(Of String, String())
        dicThermalConductivity.Add("W__m_K", {"watt/m-K", "1.0", "0.0"})
        dicThermalConductivity.Add("W__m_C", {"watt/m-C", "1.0", "0.0"})
        dicThermalConductivity.Add("kcal__hr_m_C", {"kilocalorie/hr-m-C", "1.1627782", "0.0"})
        dicThermalConductivity.Add("Btu__hr_ft_F", {"btu/hr-ft-F", "1.730643", "0.0"})
        dicThermalConductivity.Add("cal__s_cm_C", {"calorie/s-cm-C", "418.4", "0.0"})
        dicThermalConductivity.Add("Btu_in__hr_ft2_F", {"btu-in/hr-ft2-F", "0.1441314", "0.0"})
        Return dicThermalConductivity
    End Function

    Private Function GetHeatCapacity() As Dictionary(Of String, String())
        Dim dicHeatCapacity As New Dictionary(Of String, String())
        dicHeatCapacity.Add("J__kg_C", {"joule/kg-C", "1.0", "0.0"})
        dicHeatCapacity.Add("Btu__lb_F", {"btu/lb-F", "4184.0", "0.0"})
        dicHeatCapacity.Add("kcal__kg_C", {"kcal/kg-C", "4184.0", "0.0"})
        dicHeatCapacity.Add("kJ__kg_C", {"kilojoule/kg-C", "1000.0", "0.0"})
        dicHeatCapacity.Add("cal__g_C", {"cal/g-C", "4184.0", "0.0"})
        dicHeatCapacity.Add("J__kg_K", {"joule/kg-K", "1.0", "0.0"})
        Return dicHeatCapacity
    End Function

    Private Function GetSurfaceTension() As Dictionary(Of String, String())
        Dim dicSurfaceTension As New Dictionary(Of String, String())
        dicSurfaceTension.Add("N__m", {"newton/meter", "1.0", "0.0"})
        dicSurfaceTension.Add("mN__m", {"milli-newton/meter", "1.0e-3", "0.0"})
        dicSurfaceTension.Add("dyne__cm", {"dyne/centimeter", "1.0e-3", "0.0"})
        dicSurfaceTension.Add("lbf__in", {"lbf/in", "1.7513e+2", "0.0"})
        dicSurfaceTension.Add("kgf__m", {"kgf/m", "9.8067", "0.0"})
        dicSurfaceTension.Add("pdl__in", {"poundal/inch", "5.443108", "0.0"})
        Return dicSurfaceTension
    End Function

    Private Function GetDynamicViscosity() As Dictionary(Of String, String())
        Dim dicDynamicViscosity As New Dictionary(Of String, String())
        dicDynamicViscosity.Add("Pa_s", {"pascal second", "1.0", "0.0"})
        dicDynamicViscosity.Add("mPa_s", {"milli pascal second", "0.001", "0.0"})
        dicDynamicViscosity.Add("cP", {"centipoise", "0.001", "0.0"})
        dicDynamicViscosity.Add("P", {"poise", "0.1", "0.0"})
        dicDynamicViscosity.Add("lb__ft_s", {"lb/ft-sec", "1.4882", "0.0"})
        dicDynamicViscosity.Add("kgf_s__m2", {"kgf-sec/m2", "9.80665", "0.0"})
        dicDynamicViscosity.Add("lbf_s__ft2", {"lbf-sec/ft2", "47.88", "0.0"})
        dicDynamicViscosity.Add("mN_s__m2", {"mN s/m2", "0.001", "0.0"})
        Return dicDynamicViscosity
    End Function

    Private Function GetKinematicViscosity() As Dictionary(Of String, String())
        Dim dicKinematicViscosity As New Dictionary(Of String, String())
        dicKinematicViscosity.Add("cst", {"centistokes", "1.0", "0.0"})
        dicKinematicViscosity.Add("st", {"stokes", "100.0", "0.0"})
        dicKinematicViscosity.Add("m2__s", {"square meter per second", "1.0E+6", "0.0"})
        dicKinematicViscosity.Add("cm2__s", {"square centimeter per second", "1.0E+2", "0.0"})
        dicKinematicViscosity.Add("mm2__s", {"square millimeter per second", "1.0", "0.0"})
        dicKinematicViscosity.Add("in2__s", {"square inch per second", "645.16", "0.0"})
        Return dicKinematicViscosity
    End Function

    Public Enum UnitSet As Integer
        Temperature = 0
        Pressure = 1
        Mass = 2
        Length = 3
        Time = 4
        Angle = 5
        Area = 6
        Volume = 7
        Density = 8
        Force = 9
        Energy = 10
        Work = 11
        ThermalConductivity = 12
        HeatCapacity = 13
        SurfaceTension = 14
        DynamicViscosity = 15
        KinematicViscosity = 16
    End Enum

    Private Function GetDict(ByRef unitset As UnitSet) As Dictionary(Of String, String())
        Select Case unitset
            Case UnitSet.Temperature
                Return GetTemperature()
            Case UnitSet.Pressure
                Return GetPressure()
            Case UnitSet.Mass
                Return GetMass()
            Case UnitSet.Length
                Return GetLength()
            Case UnitSet.Time
                Return GetTime()
            Case UnitSet.Angle
                Return GetAngle()
            Case UnitSet.Area
                Return GetArea()
            Case UnitSet.Volume
                Return GetVolume()
            Case UnitSet.Density
                Return GetDensity()
            Case UnitSet.Force
                Return GetForce()
            Case UnitSet.Energy
                Return GetEnergy()
            Case UnitSet.Work
                Return GetWork()
            Case UnitSet.ThermalConductivity
                Return GetThermalConductivity()
            Case UnitSet.HeatCapacity
                Return GetHeatCapacity()
            Case UnitSet.SurfaceTension
                Return GetSurfaceTension()
            Case UnitSet.DynamicViscosity
                Return GetDynamicViscosity()
            Case UnitSet.KinematicViscosity
                Return GetKinematicViscosity()
        End Select
    End Function

End Class
