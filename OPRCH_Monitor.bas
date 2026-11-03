Attribute VB_Name = "OPRCH_Monitor"
Option Explicit

' ==========================================================
' OPRCH monitoring  (quantitative + qualitative)
' Per-generator:  data sheet <Ст>_<Ген>        +  chart sheet <Ст>_<Ген>_Граф
' Per-station:    data sheet <Ст>_Сумма[_<Пп>] +  chart sheet <...>_Граф
' Extra sheets:   Summary, Log
'
' Версия: 1.7.0
' Нормативная база: Методические указания по оценке участия генерирующего
'   оборудования в ОПРЧ (Приложение 14 к Техтребованиям, ред. 01.03.2026):
'   п. 4.1.2/4.1.3 (применимость количественной/качественной оценки),
'   п. 4.2.1-4.2.3 (формулы dPп, Pтп, dFр), 4.2.5 (выбор tр в пользу момента,
'   фиксирующего несоответствие), 4.2.6 (усреднение P/f, колебания >5 периодов),
'   4.2.7-4.2.8 (допуск 1 %Pном), 4.2.9 (квазиустановившиеся интервалы),
'   4.2.10 (оценка неблочной части по сумме станции -> индивидуальная оценка
'   только при нарушении), 4.2.11 (наличие реакции РЧВ турбины через 15 с),
'   4.2.14.1 (дефицит хотя бы в один из периодов), 4.2.14.4 (колебания),
'   4.4.1 (типы реакции 1-8), 4.4.3/4.4.4 (отсутствие «адекватной» реакции
'   без РМ / с РМ), Приложение (Kд(t): ТЭС 0.5 до 15 с -> 1.0 к 300/360/420 с;
'   ПГУ 0.25 до 15 с -> 0.5 к 30 с -> 1.0 к 120 с; квазиустановившийся режим:
'   |f - fср(±15 с)| <= 10 мГц; Pтреб не более 10 %Pном; идентификация
'   переходной функции h(15 с), п. 6.2 - справочно).
' Режим работы (Config W13): «Авария» - мониторинг реальных событий: для
'   оборудования без регулятора мощности (Config U = 0) принимаются S = 6 %,
'   fнч = 0.15 Гц независимо от Config (максимально допустимые по Техтребованиям);
'   «Испытание» - контрольные испытания: используются фактические S/fнч из Config,
'   проверка применимости по п. 4.1.2/4.1.3 не выполняется.
' Основные критерии:
'   - Количественный (|Pфакт|/|Pтреб|) с допуском и направлением;
'   - Перерегулирование (>100%+допуск);
'   - Качественные подкритерии: t5, t10, установившееся (среднее по хвосту);
'   - Характер переходного процесса: Монотонный / Апериодический / Колебательный;
'   - Амплитуда возмущения в %Pном и метка масштаба события;
'   - Пресеты параметров по типу оборудования (ПТУ_блок / ПТУ_неблок /
'     ГТУ / ПГУ_утил / ПГУ_сбросн / ГПА) подставляются при пустых ячейках;
'   - Учет Pmax/Pmin (диапазон регулирования): капинг Pтреб по располагаемому
'     резерву, статус 'Ограничен Pmax/Pmin' в Summary, WARN при резерве <5 %Pном,
'     горизонтальные уровни Pmax/Pmin и цветная зона за лимитом на графике.
'     Для станционных сумм Pmax_сум/Pmin_сум = сумма по включенным генераторам
'     (по паропроводам У/Э для ТЭЦ СЛПК).
'   - Реакция РЧВ через 15 с (п. 4.2.11): отдельный статус в Summary;
'   - Дефицит первичной мощности по точкам окна (п. 4.2.14.1): макс. %Pном и длительность;
'   - Вердикт по сумме станции (п. 4.2.10) со строкой в Summary и выводом о
'     необходимости индивидуальной оценки генераторов.
'   - 1.7.0: режим «Авария/Испытание»; признак РМ и компоновка (блок/неблок);
'     применимость оценки (п. 4.1.2/4.1.3); несколько событий в одной выборке
'     (листы с суффиксом _E2, _E3...); выбор tр как худшей квазиустановившейся
'     точки (п. 4.2.5/4.2.9, Приложение); колебания мощности (п. 4.2.14.4);
'     контроль качества телеизмерений; тип реакции по п. 4.4.1; оценка h(15 с)
'     идентификацией переходной функции (Приложение п. 6.2, справочно); лист
'     «Протокол»; экспорт отчёта в PDF; нестрогий поиск колонок RawData;
'     оформление графиков по И.2 (сетка не реже 15 с, подписи осей).
' ==========================================================

Public Const OPRCH_VERSION As String = "1.7.0"

Private Const SH_RAW As String = "RawData"
Private Const SH_CFG As String = "Config"
Private Const SH_SUM As String = "Summary"
Private Const SH_LOG As String = "Log"
Private Const SH_PROT As String = "Протокол"
Private Const CHART_SUFFIX As String = "_Граф"

' Нормативные константы МУ (Приложение 14)
Private Const QUANT_DF_HZ As Double = 0.2          ' п. 4.1.2: колич. оценка при |dF| >= 0.2 Гц
Private Const QUANT_MIN_SEC_TES As Double = 30#    ' п. 4.1.2: длительность >= 30 с (ТЭС)
Private Const QUANT_MIN_SEC_GPA As Double = 60#    ' п. 4.1.2: длительность >= 1 мин (ГПА/ДГУ/ГЭС)
Private Const QUAL_JUMP_SEC As Double = 10#        ' п. 4.1.3: резкое отклонение - за время <= 10 с
Private Const QUAL_MIN_SEC As Double = 20#         ' п. 4.1.3: продолжительность >= 20 с
Private Const QUAL_RM_DF_HZ As Double = 0.15       ' п. 4.4.4: с РМ - выход за 50±0.15 Гц
Private Const QUAL_NORM_DF_MAX_HZ As Double = 0.4  ' п. 4.4.3: без РМ - скачок ±0.2..0.4 Гц
Private Const NO_RM_S_PCT As Double = 6#           ' без РМ в режиме «Авария»: S = 6 %
Private Const NO_RM_FNCH_HZ As Double = 0.15       ' без РМ в режиме «Авария»: fнч = 0.15 Гц
Private Const QS_HALF_WINDOW_SEC As Double = 15#   ' Приложение: квазиуст. режим - отрезок ±15 с
Private Const QS_TOL_HZ As Double = 0.01           ' Приложение: |f - fср| <= 10 мГц
Private Const OSC_MIN_PERIODS As Long = 5          ' п. 4.2.6/4.2.14.4: колебания > 5 периодов
Private Const OSC_MIN_AMPL_PCT As Double = 1#      ' п. 4.2.14.4: амплитуда > 1 %Pном
Private Const TELEM_MAX_REPEAT As Long = 11        ' Прил. 1 п. 6.1.8.3: не более 11 одинаковых подряд
Private Const H15_REF As Double = 0.3              ' ориентир для h(15 с) (Приложение п. 6.2, справочно)
Private Const EVENT_MIN_OUTSIDE_SEC As Double = 5# ' фильтр коротких выбросов частоты при поиске событий

Private m_LogRow As Long
Private m_KdProfiles As Object   ' key=EQUIPTYPE, value=Array(t0,kd0,t1,kd1,t2,kd2,t3,kd3)
Private m_ChartOpts As Object    ' key=REL_*/ABS_*, value=Boolean
Private m_Protocol As Collection ' строки текстового протокола (лист «Протокол»)

Private Type TSettings
    FNom As Double
    EventStart As Double
    AutoStart As Boolean
    QuantIntervalSec As Double
    QuantTolPct As Double
    WorkThresholdMW As Double
    PreBufferSec As Double
    ChartIntervalSec As Double
    SteadyWindowSec As Double
    React15ThrPct As Double   ' порог наличия реакции РЧВ через 15 с, %Pном (п. 4.2.11)
    AvgWindowSec As Double    ' окно усреднения P0/Pтек, с (п. 4.2.6); 0 = текущие значения
    Mode As String            ' «Авария» / «Испытание»
    IsEmergency As Boolean    ' True = «Авария»
    MaxEventSec As Double     ' макс. длительность окна одного события, с
    MaxEvents As Long         ' макс. число событий в выборке (1 = только первое)
    MaxReqPct As Double       ' ограничение Pтреб, %Pном (Приложение: 10 %; 0 = без ограничения)
End Type

' Событие (интервал выхода частоты за fнч)
Private Type TEvent
    StartRow As Long          ' последняя строка до выхода за fнч (t0)
    FirstExceedRow As Long    ' первая строка за fнч
    EndRow As Long            ' конец окна события
End Type

Private Type TGenCfg
    Station As String
    Generator As String
    PowerHeader As String
    FreqHeader As String
    EquipType As String
    PNom As Double
    SPct As Double
    Fnch As Double
    Enabled As Boolean
    QualEnabled As Boolean
    T5Sec As Double
    Dp5Pct As Double
    T10Sec As Double
    Dp10Pct As Double
    SteadyTolPct As Double
    InStationSum As Boolean
    CheckSteady As Boolean
    Paroprovod As String
    PMax As Double          ' максимальная эксплуатационная мощность, МВт
    PMin As Double          ' технический минимум, МВт
    HasRM As Boolean        ' оснащён регулятором мощности (Config U; пусто = 1)
    Layout As String        ' «блок» / «неблок» (Config V; пусто = по типу оборудования)
    SPctCfg As Double       ' S из Config (до подстановки 6 % для оборудования без РМ)
    FnchCfg As Double       ' fнч из Config (до подстановки 0.15 Гц)
    ForcedNoRM As Boolean   ' True, если в режиме «Авария» применены S = 6 %, fнч = 0.15 Гц
End Type

Private Type TGenResult
    StartRow As Long
    EndQuantRow As Long
    EndQualRow As Long
    StartTime As Variant
    FirstExceedTime As Variant
    P0 As Double
    PTek As Double
    PsteadyAvg As Double
    Df As Double
    Dfr As Double
    PReq As Double
    PReqOrig As Double      ' исходный требуемый, до ограничения Pmax/Pmin
    PFact As Double
    AmplPctPnom As Double
    AmplitudeTag As String
    QuantPct As Double
    QuantPass As Boolean
    Overshoot As Boolean
    TransientType As String
    NumExtrema As Long
    PReqSteady As Double
    QualPass As Boolean
    QualT5Pass As Boolean
    QualT10Pass As Boolean
    QualSteadyPass As Boolean
    T5FactSec As Double
    T10FactSec As Double
    QualFailedList As String
    QualReason As String
    PMaxEff As Double       ' эффективная Pmax (g.PMax или g.PNom по умолчанию)
    PMinEff As Double       ' эффективная Pmin (g.PMin или 0)
    ReservePlus As Double   ' располагаемый резерв '+' = max(0, PMax - P0)
    ReserveMinus As Double  ' располагаемый резерв '-' = max(0, P0 - PMin)
    Limited As Boolean      ' Pтреб был ограничен диапазоном
    LimitType As String     ' 'Pmax' или 'Pmin'
    KdUsedQuant As Double   ' Kд, принятый в количественном расчёте
    KdProfile As String     ' справка по применённому профилю Kд(t)
    PReqStatic As Double    ' Pтреб при Kд=1 (статическая характеристика), с учётом резерва
    React15Status As String ' "ОК" / "Нарушение" / "Н/П" (не применимо) - п. 4.2.11
    React15Pass As Boolean
    DP15 As Double          ' изменение мощности через 15 с после выхода за fнч, МВт
    React15Note As String
    DeficitMaxPct As Double ' макс. дефицит первичной мощности в окне, %Pном (п. 4.2.14.1)
    DeficitSec As Double    ' суммарная длительность дефицита, с
    Verdict As String       ' итоговый вывод по генератору
    ' --- 1.7.0 ---
    EventIdx As Long        ' номер события в выборке (1..N)
    EventEndRow As Long     ' конец окна события
    QuantApplicable As Boolean  ' п. 4.1.2
    QualApplicable As Boolean   ' п. 4.1.3 / 4.4.3 / 4.4.4
    Applicability As String     ' текст статуса применимости
    DfMaxEvent As Double        ' макс. отклонение частоты за событие (со знаком), Гц
    DurAbove02Sec As Double     ' длительность |dF| >= 0.2 Гц (непрерывная макс.), с
    JumpSec As Double           ' время нарастания отклонения до максимума, с
    DurOutsideSec As Double     ' длительность |dF| > fнч, с
    TrRow As Long               ' строка tр (п. 4.2.5: худшая квазиустановившаяся точка)
    TrTime As Variant
    TrSec As Double             ' tр - t0, с
    TrQuasiSteady As Boolean    ' tр в квазиустановившемся режиме по частоте
    PReqRef As Double           ' Pтреб на контрольном интервале (Config W5), МВт
    PFactRef As Double          ' Pфакт на контрольном интервале, МВт
    OscViolation As Boolean     ' п. 4.2.14.4
    OscPeriods As Long
    OscAmplPct As Double
    OscNote As String
    TelemetryOK As Boolean
    TelemetryNote As String
    ReactionType As String      ' п. 4.4.1: «1 - адекватная» ... «8 - телеизмерение»
    ReactionCode As Long
    QualClauseVerdict As String ' вывод по п. 4.4.3/4.4.4
    H15 As Double               ' оценка переходной функции h(15 с); -999 = не вычислено
    H15Note As String
End Type

' Результат оценки по сумме станции (п. 4.2.10)
Private Type TStationResult
    Station As String
    Paroprovod As String
    Count As Long
    StartTime As Variant
    FirstExceedTime As Variant
    P0 As Double
    PTek As Double
    Df As Double
    PReq As Double
    PReqOrig As Double
    PFact As Double
    QuantPct As Double
    QuantPass As Boolean
    Limited As Boolean
    LimitType As String
    PMaxSum As Double
    PMinSum As Double
    ReservePlus As Double
    ReserveMinus As Double
    PNomSum As Double
    React15Status As String
    React15Pass As Boolean
    DP15 As Double
    React15Note As String
    DeficitMaxPct As Double
    DeficitSec As Double
    Verdict As String
    SheetName As String
    ChartSheetName As String
    EventIdx As Long
    IsNeblok As Boolean         ' в сумме есть генераторы неблочной компоновки (п. 4.2.10 применим)
    QuantApplicable As Boolean
    Applicability As String
    TrTime As Variant
    TrSec As Double
    TrQuasiSteady As Boolean
End Type

' ==========================================================
' Точки входа
' ==========================================================

Public Sub SetupOPRCHTemplate()
    Dim wsRaw As Worksheet, wsCfg As Worksheet, wsSum As Worksheet

    Set wsRaw = EnsureSheet(SH_RAW)
    Set wsCfg = EnsureSheet(SH_CFG)
    Set wsSum = EnsureSheet(SH_SUM)

    wsRaw.Cells.Clear
    wsCfg.Cells.Clear
    wsSum.Cells.Clear

    wsRaw.Range("A1").Value = "Время"
    wsRaw.Range("B1").Value = "Частота"
    wsRaw.Range("C1").Value = "ТГ-1"
    wsRaw.Range("D1").Value = "ТГ-2"
    wsRaw.Range("E1").Value = "ТГ-3"

    wsCfg.Range("A1:V1").Value = Array( _
        "Станция", "Генератор", "Колонка_мощности", "Колонка_частоты", "Тип_оборудования", _
        "Pном, МВт", "S, %", "fнч, Гц", "Вкл (1/0)", "Кач_вкл (1/0)", _
        "t5, c", "dP5, %Pном", "t10, c", "dP10, %Pном", "Уст_допуск, %Pном", _
        "В сумму станции (1/0)", "Контр_уст (1/0)", "Паропровод", _
        "Pmax, МВт", "Pmin, МВт", "РМ (1/0)", "Компоновка (блок/неблок)" _
    )

    wsCfg.Cells(2, 1).Resize(1, 22).Value = Array("Сосногорская ТЭЦ", "ТГ-5", "ТГ-5", "Частота", "ПТУ_неблок", 55, 4.2, 0.105, 1, 1, 15, 5, 420, 10, 1, 1, 1, "", 55, 15, 1, "неблок")
    wsCfg.Cells(3, 1).Resize(1, 22).Value = Array("Сосногорская ТЭЦ", "ТГ-7", "ТГ-7", "Частота", "ПТУ_неблок", 60, 4.5, 0.11, 1, 1, 15, 5, 420, 10, 1, 1, 1, "", 60, 20, 1, "неблок")
    wsCfg.Cells(4, 1).Resize(1, 22).Value = Array("ТЭЦ СЛПК", "ТГ-2Э", "ТГ-2Э", "Частота", "ПТУ_неблок", 50, 4.5, 0.15, 1, 1, 15, 5, 420, 10, 1, 1, 1, "Э", 50, 15, 1, "неблок")
    wsCfg.Cells(5, 1).Resize(1, 22).Value = Array("ТЭЦ СЛПК", "ТГ-5У", "ТГ-5У", "Частота", "ПТУ_неблок", 87.7, 4.2, 0.15, 1, 1, 15, 5, 420, 10, 1, 1, 1, "У", 87.7, 25, 1, "неблок")

    wsCfg.Range("W1").Value = "Глобальные настройки"
    wsCfg.Cells(2, 23).Resize(1, 2).Value = Array("fном, Гц", 50)
    wsCfg.Cells(3, 23).Resize(1, 2).Value = Array("Время начала события", "")
    wsCfg.Cells(4, 23).Resize(1, 2).Value = Array("Автопоиск старта (1/0)", 1)
    wsCfg.Cells(5, 23).Resize(1, 2).Value = Array("Колич. интервал, с", 82)
    wsCfg.Cells(6, 23).Resize(1, 2).Value = Array("Допуск количеств., %", 10)
    wsCfg.Cells(7, 23).Resize(1, 2).Value = Array("Порог включения в работу, МВт", 1)
    wsCfg.Cells(8, 23).Resize(1, 2).Value = Array("Pre-start буфер, с", 5)
    wsCfg.Cells(9, 23).Resize(1, 2).Value = Array("Интервал графика, с", 120)
    wsCfg.Cells(10, 23).Resize(1, 2).Value = Array("Окно установив., с", 30)
    wsCfg.Cells(11, 23).Resize(1, 2).Value = Array("Порог реакции 15 с, %Pном", 0.5)
    wsCfg.Cells(12, 23).Resize(1, 2).Value = Array("Усреднение P0/Pтек, с (0=текущие)", 0)
    wsCfg.Cells(13, 23).Resize(1, 2).Value = Array("Режим (Авария/Испытание)", "Авария")
    wsCfg.Cells(14, 23).Resize(1, 2).Value = Array("Макс. окно события, с", 600)
    wsCfg.Cells(15, 23).Resize(1, 2).Value = Array("Макс. число событий (1=только первое)", 3)
    wsCfg.Cells(16, 23).Resize(1, 2).Value = Array("Ограничение Pтреб, %Pном (0=нет)", 10)
    wsCfg.Range("X13").NumberFormat = "@"

    ' Профили Kд(t) по Приложению к МУ оценки участия в ОПРЧ (Приложение 14 к Техтребованиям):
    '   ТЭС (кроме ПГУ): Kд = 0.5 на [0;15] с, далее линейно до 1.0 к T
    '       (T = 300 с газомазутные блоки, 360 с пылеугольные, 420 с общий паропровод);
    '   ПГУ/ГТУ: Kд = 0.25 на [0;15] с, 0.5 к 30 с, 1.0 к 120 с;
    '   ГПА: нормируется интервалом 1 мин (10 %Pном к 60 с) - ориентир.
    wsCfg.Range("AA1").Value = "Профили Kд(t) (абсолютные значения Kд)"
    wsCfg.Range("AA2:AI2").Value = Array("Тип_оборудования", "t0, с", "Kд0", "t1, с", "Kд1", "t2, с", "Kд2", "t3, с", "Kд3")
    wsCfg.Cells(3, 27).Resize(1, 9).Value = Array("ПТУ_блок", 0, 0.5, 15, 0.5, 300, 1, 600, 1)
    wsCfg.Cells(4, 27).Resize(1, 9).Value = Array("ПТУ_блок_уголь", 0, 0.5, 15, 0.5, 360, 1, 600, 1)
    wsCfg.Cells(5, 27).Resize(1, 9).Value = Array("ПТУ_неблок", 0, 0.5, 15, 0.5, 420, 1, 600, 1)
    wsCfg.Cells(6, 27).Resize(1, 9).Value = Array("ГТУ", 0, 0.25, 15, 0.25, 30, 0.5, 120, 1)
    wsCfg.Cells(7, 27).Resize(1, 9).Value = Array("ПГУ_утил", 0, 0.25, 15, 0.25, 30, 0.5, 120, 1)
    wsCfg.Cells(8, 27).Resize(1, 9).Value = Array("ПГУ_сбросн", 0, 0.5, 15, 0.5, 300, 1, 600, 1)
    wsCfg.Cells(9, 27).Resize(1, 9).Value = Array("ГПА", 0, 0.5, 15, 0.5, 60, 1, 600, 1)

    wsCfg.Range("AJ1").Value = "Отображение рядов графиков (1/0)"
    wsCfg.Range("AJ2:AL2").Value = Array("Ряд", "Относит.", "Абсолют.")
    wsCfg.Range("AJ3:AJ10").Value = WorksheetFunction.Transpose(Array( _
        "Pфакт", "Pтреб", "+Допуск", "-Допуск", "Pmax/Pmin", "Частота", "Маркеры t5/t10/fнч", "Резерв" _
    ))
    wsCfg.Range("AK3:AL10").Value = 1
    wsCfg.Range("AK10:AL10").Value = 0

    wsCfg.Columns("A:AL").AutoFit
    wsRaw.Columns("A:E").AutoFit
    EnsureControlButtons wsCfg
    EnsureChartOptionCheckboxes wsCfg

    MsgBox "Шаблон создан. Заполните RawData/Config и нажмите кнопку 'Запустить мониторинг ОПРЧ'.", vbInformation
End Sub

' Экспорт отчёта (Summary, Протокол, Log и листы графиков) в PDF рядом с книгой
Public Sub ExportOPRCHReport()
    On Error GoTo EH
    Dim ws As Worksheet, names As Collection, i As Long
    Dim arr() As String, pdfPath As String, baseName As String
    Set names = New Collection
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SH_SUM Or ws.Name = SH_PROT Or ws.Name = SH_LOG Then
            names.Add ws.Name
        ElseIf Right$(ws.Name, Len(CHART_SUFFIX)) = CHART_SUFFIX Or Right$(ws.Name, 3) = "_Гр" Then
            names.Add ws.Name
        End If
    Next ws
    If names.Count = 0 Then
        MsgBox "Нет листов для экспорта. Сначала запустите мониторинг.", vbExclamation
        Exit Sub
    End If
    ReDim arr(0 To names.Count - 1)
    For i = 1 To names.Count
        arr(i - 1) = CStr(names(i))
    Next i
    baseName = ThisWorkbook.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Сначала сохраните книгу - PDF создаётся в той же папке.", vbExclamation
        Exit Sub
    End If
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & baseName & "_Отчет_ОПРЧ_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"
    For i = 0 To UBound(arr)
        With ThisWorkbook.Worksheets(arr(i)).PageSetup
            On Error Resume Next
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = False
            On Error GoTo EH
        End With
    Next i
    ThisWorkbook.Worksheets(arr).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
                                    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ThisWorkbook.Worksheets(SH_SUM).Select
    MsgBox "Отчёт сохранён: " & pdfPath, vbInformation
    Exit Sub
EH:
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SUM).Select
    MsgBox "Ошибка экспорта в PDF: " & Err.Description, vbCritical
End Sub

Public Sub AnalyzeOPRCH()
    On Error GoTo EH

    Dim wsRaw As Worksheet, wsCfg As Worksheet, wsSummary As Worksheet
    Dim st As TSettings
    Dim timeCol As Long, cfgLast As Long, r As Long, outRow As Long
    Dim g As TGenCfg, res As TGenResult
    Dim targetSheets As Collection
    Dim stepName As String
    Dim t0Run As Double
    Dim evList() As TEvent, nEv As Long, k As Long
    Dim fCol As Long

    t0Run = Timer
    stepName = "Подготовка листов"
    Set wsRaw = GetRequiredSheet(SH_RAW)
    Set wsCfg = GetRequiredSheet(SH_CFG)
    Set wsSummary = EnsureSheet(SH_SUM)

    stepName = "Чтение настроек"
    st = ReadSettings(wsCfg)
    LoadKdProfiles wsCfg
    LoadChartDisplayOptions wsCfg
    EnsureChartOptionCheckboxes wsCfg
    timeCol = FindHeaderCol(wsRaw, "Время")
    If timeCol = 0 Then Err.Raise vbObjectError + 2001, , "В RawData не найдена колонка 'Время'."

    cfgLast = LastUsedRow(wsCfg)
    If cfgLast < 2 Then Err.Raise vbObjectError + 2002, , "В Config нет строк генераторов."

    stepName = "Подготовка лога"
    InitLog
    Set m_Protocol = New Collection
    AppendLog "INFO", "Settings", "Режим: " & st.Mode & IIf(st.IsEmergency, _
        " (для оборудования без РМ принимаются S = " & NO_RM_S_PCT & " %, fнч = " & NO_RM_FNCH_HZ & " Гц)", _
        " (используются фактические S/fнч из Config; применимость по п. 4.1.2/4.1.3 не проверяется)")

    stepName = "Валидация Config/RawData"
    ValidateInputs wsRaw, wsCfg, st, timeCol

    stepName = "Очистка выходных листов"
    Set targetSheets = New Collection
    CollectOldOutputSheets targetSheets
    DeleteOutputSheets targetSheets

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    stepName = "Подготовка Summary"
    wsSummary.Cells.Clear
    wsSummary.Range("A1:AZ1").Value = Array( _
        "Станция", "Генератор", "Тип", _
        "Старт (расч.)", "Время выхода за fнч", _
        "P0, МВт", "Pтек(tр), МВт", "Pуст_сред, МВт", _
        "dF, Гц", "dFr(tр), Гц", _
        "Pтреб(tр), МВт", "Pфакт(tр), МВт", _
        "Амплитуда, %Pном", "Масштаб события", _
        "Колич. %", "Колич. статус", "Перерегулирование", _
        "Характер процесса", "Экстремумов", _
        "Кач. статус", "Кач.t5", "Кач.t10", "Кач.уст", _
        "Проваленные подпункты", "t5 факт, c", "t10 факт, c", _
        "Лист", "Лист графика", _
        "Pmax, МВт", "Pmin, МВт", "Резерв '+', МВт", "Резерв '-', МВт", _
        "Pтреб исх., МВт", "Ограничение", _
        "Примечание", _
        "Реакция 15 с", "dP15, МВт", "Дефицит макс, %Pном", "Длит. дефицита, с", _
        "Вывод", _
        "Событие", "Применимость (п. 4.1.2/4.1.3)", "tр", "tр, с от старта", "Режим в tр", _
        "Pтреб(Tинт), МВт", "Pфакт(Tинт), МВт", _
        "Колебания (п. 4.2.14.4)", "Качество ТИ", "Тип реакции (п. 4.4.1)", _
        "h(15 с), Прил. п. 6.2", "Кач. вывод (п. 4.4.3/4.4.4)" _
    )
    outRow = 2

    For r = 2 To cfgLast
        stepName = "Чтение Config, строка " & r
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NextGen

        g = ReadGenCfgEff(wsCfg, r, st)
        If Not g.Enabled Then GoTo NextGen

        If Not ValidateGenCfg(g) Then
            WriteSummaryInvalid wsSummary, outRow, g
            AppendLog "Config", g.Station & "/" & g.Generator, "Пропущен: не заполнены обязательные параметры"
            outRow = outRow + 1
            GoTo NextGen
        End If

        stepName = "Поиск событий для " & g.Generator
        fCol = FindHeaderColFuzzy(wsRaw, g.FreqHeader)
        If fCol = 0 Then Err.Raise vbObjectError + 2102, , "Не найдена колонка частоты '" & g.FreqHeader & "' для " & g.Generator
        nEv = FindEvents(wsRaw, timeCol, fCol, st, g.Fnch, evList)
        If nEv > 1 Then AppendLog "INFO", g.Station & "/" & g.Generator, "Найдено событий (выходов за fнч): " & nEv

        For k = 1 To nEv
            stepName = "Расчет генератора " & g.Generator & " (событие " & k & ")"
            res = AnalyzeOneGenerator(wsRaw, st, g, evList(k), k)
            stepName = "Запись листа данных " & g.Generator
            WriteGeneratorSheet wsRaw, st, g, res
            stepName = "Запись графика " & g.Generator
            WriteGeneratorChartSheet st, g, res
            stepName = "Запись Summary для " & g.Generator
            WriteSummaryRow wsSummary, outRow, g, res
            AddProtocolGenerator g, res
            outRow = outRow + 1
        Next k

NextGen:
    Next r

    stepName = "Расчет суммарных листов станций"
    BuildStationAggregates wsRaw, wsCfg, wsSummary, st, outRow

    stepName = "Оформление Summary"
    wsSummary.Columns("A:AZ").AutoFit
    wsSummary.Columns("AN").ColumnWidth = 90
    wsSummary.Columns("AZ").ColumnWidth = 70
    wsSummary.Columns("AI").ColumnWidth = 60
    wsSummary.Columns("AP").ColumnWidth = 45
    wsSummary.Range("D:E").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsSummary.Range("AQ:AQ").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ApplySummaryConditionalFormat wsSummary
    WriteVersionStamp wsSummary, wsRaw, t0Run

    stepName = "Формирование протокола"
    WriteProtocolSheet st, wsRaw

    stepName = "Финализация лога и кнопок"
    FinalizeLog
    EnsureControlButtons wsCfg

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Мониторинг ОПРЧ завершен (v" & OPRCH_VERSION & "). Время: " _
        & Format(Timer - t0Run, "0.0") & " c.", vbInformation
    Exit Sub

EH:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Ошибка AnalyzeOPRCH (" & stepName & "): " & Err.Description, vbCritical
End Sub

Public Sub ClearOPRCHResults()
    Dim targetSheets As Collection
    Set targetSheets = New Collection
    CollectOldOutputSheets targetSheets
    Application.DisplayAlerts = False
    DeleteOutputSheets targetSheets
    Application.DisplayAlerts = True
    MsgBox "Результаты очищены (" & targetSheets.Count & " вкладок).", vbInformation
End Sub

Public Sub ApplyPresetsToConfig()
    Dim wsCfg As Worksheet, r As Long, cfgLast As Long, changed As Long
    Dim et As String, g As TGenCfg, pr As TGenCfg
    Set wsCfg = GetRequiredSheet(SH_CFG)
    cfgLast = LastUsedRow(wsCfg)
    changed = 0
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NX
        et = Trim$(CStr(wsCfg.Cells(r, 5).Value))
        If Len(et) = 0 Then GoTo NX
        pr = GetPreset(et)
        If NzD(wsCfg.Cells(r, 11).Value, 0) <= 0 Then wsCfg.Cells(r, 11).Value = pr.T5Sec: changed = changed + 1
        If NzD(wsCfg.Cells(r, 12).Value, 0) <= 0 Then wsCfg.Cells(r, 12).Value = pr.Dp5Pct: changed = changed + 1
        If NzD(wsCfg.Cells(r, 13).Value, 0) <= 0 Then wsCfg.Cells(r, 13).Value = pr.T10Sec: changed = changed + 1
        If NzD(wsCfg.Cells(r, 14).Value, 0) <= 0 Then wsCfg.Cells(r, 14).Value = pr.Dp10Pct: changed = changed + 1
        If NzD(wsCfg.Cells(r, 15).Value, 0) <= 0 Then wsCfg.Cells(r, 15).Value = pr.SteadyTolPct: changed = changed + 1
        If NzD(wsCfg.Cells(r, 8).Value, -1) < 0 Then wsCfg.Cells(r, 8).Value = pr.Fnch: changed = changed + 1
        If Len(Trim$(CStr(wsCfg.Cells(r, 21).Value))) = 0 Then wsCfg.Cells(r, 21).Value = 1: changed = changed + 1
        If Len(Trim$(CStr(wsCfg.Cells(r, 22).Value))) = 0 Then
            wsCfg.Cells(r, 22).Value = IIf(InStr(UCase$(et), "НЕБЛОК") > 0, "неблок", "блок"): changed = changed + 1
        End If
NX:
    Next r
    MsgBox "Пресеты применены. Заполнено ячеек: " & changed, vbInformation
End Sub

' ==========================================================
' Анализ одного генератора
' ==========================================================

Private Function AnalyzeOneGenerator(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, _
                                     ByRef ev As TEvent, ByVal evIdx As Long) As TGenResult
    On Error GoTo EH

    Dim res As TGenResult
    Dim timeCol As Long, pCol As Long, fCol As Long
    Dim startRow As Long, endQ As Long, endQual As Long, firstExceedRow As Long, evEnd As Long
    Dim p0 As Double, ptek As Double, df As Double, dfr As Double, preq As Double, pfact As Double
    Dim kdQuant As Double, tQuantSec As Double
    Dim qpct As Double, qpass As Boolean
    Dim calcStep As String

    calcStep = "Поиск колонок"
    timeCol = FindHeaderCol(wsRaw, "Время")
    pCol = FindHeaderColFuzzy(wsRaw, g.PowerHeader)
    fCol = FindHeaderColFuzzy(wsRaw, g.FreqHeader)
    If pCol = 0 Then Err.Raise vbObjectError + 2101, , "Не найдена колонка мощности '" & g.PowerHeader & "' для " & g.Generator
    If fCol = 0 Then Err.Raise vbObjectError + 2102, , "Не найдена колонка частоты '" & g.FreqHeader & "' для " & g.Generator

    calcStep = "Границы события"
    startRow = ev.StartRow
    firstExceedRow = ev.FirstExceedRow
    evEnd = ev.EndRow
    If evEnd < startRow Then evEnd = startRow
    res.EventIdx = evIdx
    res.EventEndRow = evEnd
    calcStep = "Определение конца контрольного интервала"
    endQ = RowByTimeOffset(wsRaw, timeCol, startRow, st.QuantIntervalSec)
    calcStep = "Определение конца качественного интервала"
    endQual = RowByTimeOffset(wsRaw, timeCol, startRow, g.T10Sec)

    calcStep = "Чтение P0"
    ' п. 4.2.6: значения P вычисляются как средние на интервале вблизи момента
    ' оценки; при стабильных значениях допускается брать текущие (AvgWindowSec = 0).
    p0 = AvgPowerBackward(wsRaw, pCol, timeCol, startRow, st.AvgWindowSec)

    calcStep = "Применимость оценки (п. 4.1.2/4.1.3)"
    EvaluateApplicability wsRaw, st, g, fCol, timeCol, startRow, firstExceedRow, evEnd, res

    calcStep = "Учет Pmax/Pmin (резерв)"
    Dim pMaxEff As Double, pMinEff As Double
    Dim reservePlus As Double, reserveMinus As Double
    pMaxEff = g.PMax
    If pMaxEff <= 0 Then pMaxEff = g.PNom
    pMinEff = g.PMin
    If pMinEff < 0 Then pMinEff = 0
    reservePlus = pMaxEff - p0
    reserveMinus = p0 - pMinEff
    If reservePlus < 0 Then reservePlus = 0
    If reserveMinus < 0 Then reserveMinus = 0

    calcStep = "Контрольный интервал (Config W5): Pтреб/Pфакт"
    Dim dfRef As Double, dfrRef As Double, preqRefOrig As Double, preqRef As Double
    tQuantSec = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(endQ, timeCol).Value)
    dfRef = MaxAbsDeviationInWindow(wsRaw, fCol, startRow, endQ, st.FNom)
    dfrRef = DeadbandDeviation(dfRef, g.Fnch)
    preqRefOrig = PReqCalc(g.SPct, g.PNom, st.FNom, DynamicKdByTime(g.EquipType, tQuantSec), dfrRef, st.MaxReqPct)
    preqRef = ClampReserve(preqRefOrig, reservePlus, reserveMinus)
    res.PReqRef = preqRef
    res.PFactRef = AvgPowerBackward(wsRaw, pCol, timeCol, endQ, st.AvgWindowSec) - p0

    calcStep = "Выбор tр (п. 4.2.5/4.2.9)"
    ' tр - худшая квазиустановившаяся по частоте точка окна события (макс. дефицит);
    ' при отсутствии квазиустановившихся точек - конец контрольного интервала.
    SelectEvaluationPoint wsRaw, st, g, pCol, fCol, timeCol, startRow, firstExceedRow, evEnd, endQ, p0, reservePlus, reserveMinus, res
    endQ = res.TrRow
    tQuantSec = res.TrSec

    calcStep = "Чтение Pтек(tр)"
    ptek = AvgPowerBackward(wsRaw, pCol, timeCol, endQ, st.AvgWindowSec)
    calcStep = "Расчет dF/dFr"
    df = MaxAbsDeviationInWindow(wsRaw, fCol, startRow, evEnd, st.FNom)
    If res.TrQuasiSteady Then
        ' п. 4.2.6: в квазиустановившемся режиме - среднее значение частоты на отрезке ±15 с
        dfr = DeadbandDeviation(MeanFreqAround(wsRaw, fCol, timeCol, endQ, QS_HALF_WINDOW_SEC, st.FNom) - st.FNom, g.Fnch)
    Else
        dfr = DeadbandDeviation(MaxAbsDeviationInWindow(wsRaw, fCol, startRow, endQ, st.FNom), g.Fnch)
    End If
    If dfr = 0 And dfrRef <> 0 Then dfr = dfrRef

    calcStep = "Расчет требуемой мощности"
    Dim preqOrig As Double
    kdQuant = DynamicKdByTime(g.EquipType, tQuantSec)
    preqOrig = PReqCalc(g.SPct, g.PNom, st.FNom, kdQuant, dfr, st.MaxReqPct)

    calcStep = "Ограничение Pтреб резервом"
    ' Капинг: реально достижимое значение с учётом эксплуатационного диапазона
    preq = preqOrig
    If preq > reservePlus Then
        preq = reservePlus
        res.Limited = True
        res.LimitType = "Pmax"
    ElseIf preq < -reserveMinus Then
        preq = -reserveMinus
        res.Limited = True
        res.LimitType = "Pmin"
    Else
        res.Limited = False
        res.LimitType = ""
    End If

    calcStep = "Расчет фактической мощности"
    pfact = ptek - p0

    calcStep = "Количественный критерий"
    If dfr = 0 Then
        qpct = 100
        qpass = True
    ElseIf Abs(preq) < 0.000001 Then
        ' Pтреб_исх есть, но капинг обнулил: резерв в нужную сторону равен нулю.
        ' Ничего требовать нельзя - считаем статус ОК (участие ограничено диапазоном).
        qpct = 100
        qpass = True
    ElseIf SgnNZ(pfact) <> SgnNZ(preq) Then
        qpct = 0
        qpass = False
    Else
        qpct = 100# * SafeDiv(Abs(pfact), Abs(preq), 0)
        qpass = (qpct >= (100# - st.QuantTolPct))
    End If

    res.StartRow = startRow
    res.EndQuantRow = endQ
    res.EndQualRow = endQual
    res.StartTime = wsRaw.Cells(startRow, timeCol).Value
    If firstExceedRow > 0 Then
        res.FirstExceedTime = wsRaw.Cells(firstExceedRow, timeCol).Value
    Else
        res.FirstExceedTime = ""
    End If
    res.P0 = p0
    res.PTek = ptek
    res.Df = df
    res.Dfr = dfr
    res.PReq = preq
    res.PReqOrig = preqOrig
    res.PFact = pfact
    res.KdUsedQuant = kdQuant
    res.KdProfile = KdProfileText(g.EquipType)

    ' Статический Pтреб (Kд = 1, окончание переходного процесса), ограниченный резервом.
    ' Используется как цель качественных ступеней t5/t10 (п. 4.2.2, 4.2.7).
    calcStep = "Статический Pтреб"
    Dim preqStatic As Double
    preqStatic = ClampReserve(PReqCalc(g.SPct, g.PNom, st.FNom, 1#, dfr, st.MaxReqPct), reservePlus, reserveMinus)
    res.PReqStatic = preqStatic
    res.PMaxEff = pMaxEff
    res.PMinEff = pMinEff
    res.ReservePlus = reservePlus
    res.ReserveMinus = reserveMinus
    res.QuantPct = qpct
    res.QuantPass = qpass
    res.Overshoot = (qpass And qpct > (100# + st.QuantTolPct))

    ' WARN при недостаточном резерве в нужную сторону (порог 5 %Pном)
    calcStep = "Проверка резерва 5 %Pном"
    Dim minReservePct As Double, minReserveMW As Double, needSign As Integer
    minReservePct = 5#
    minReserveMW = minReservePct / 100# * g.PNom
    needSign = SgnNZ(preqOrig)
    If needSign = 1 And reservePlus < minReserveMW Then
        AppendLog "WARN", g.Station & "/" & g.Generator, _
                  "Резерв '+' = " & Format(reservePlus, "0.0") & " МВт < " & _
                  Format(minReserveMW, "0.0") & " МВт (5 %Pном). Возможно ограничение по Pmax."
    ElseIf needSign = -1 And reserveMinus < minReserveMW Then
        AppendLog "WARN", g.Station & "/" & g.Generator, _
                  "Резерв '-' = " & Format(reserveMinus, "0.0") & " МВт < " & _
                  Format(minReserveMW, "0.0") & " МВт (5 %Pном). Возможно ограничение по Pmin."
    End If
    If res.Limited Then
        AppendLog "INFO", g.Station & "/" & g.Generator, _
                  "Pтреб ограничен " & res.LimitType & ": было " & Format(preqOrig, "0.000") & _
                  " МВт, принято " & Format(preq, "0.000") & " МВт."
    End If

    calcStep = "Амплитуда события"
    ' Амплитуда берётся по исходному (неограниченному) Pтреб - она характеризует
    ' масштаб возмущения, а не способность генератора его отработать.
    If g.PNom > 0 Then
        res.AmplPctPnom = 100# * Abs(preqOrig) / g.PNom
    Else
        res.AmplPctPnom = 0
    End If
    res.AmplitudeTag = AmplitudeTag(res.AmplPctPnom, dfr <> 0)

    calcStep = "Реакция РЧВ через 15 с (п. 4.2.11)"
    Dim r15Status As String, r15Pass As Boolean, dp15 As Double, r15Note As String
    EvaluateReaction15 wsRaw, st, pCol, fCol, timeCol, startRow, firstExceedRow, p0, g.PNom, g.SPct, _
                       g.Fnch, g.EquipType, SgnNZ(preqOrig), reservePlus, reserveMinus, _
                       r15Status, r15Pass, dp15, r15Note
    res.React15Status = r15Status
    res.React15Pass = r15Pass
    res.DP15 = dp15
    res.React15Note = r15Note
    If r15Status = "Нарушение" Then
        AppendLog "WARN", g.Station & "/" & g.Generator, "п. 4.2.11: " & r15Note
    End If

    calcStep = "Дефицит первичной мощности по точкам окна (п. 4.2.14.1)"
    ' Окно наблюдения - всё событие (до возврата частоты в зону нечувствительности
    ' или до границы окна события), а не только контрольный интервал.
    Dim defMaxPct As Double, defSec As Double
    EvaluateDeficit wsRaw, st, pCol, fCol, timeCol, startRow, evEnd, p0, g.PNom, g.SPct, g.Fnch, _
                    g.EquipType, reservePlus, reserveMinus, defMaxPct, defSec
    res.DeficitMaxPct = defMaxPct
    res.DeficitSec = defSec

    calcStep = "Характер переходного процесса"
    EvaluateTransient wsRaw, st, g, res, pCol, timeCol

    calcStep = "Колебания мощности (п. 4.2.6, 4.2.14.4)"
    EvaluateOscillations wsRaw, st, g, res, pCol, fCol, timeCol, startRow, evEnd

    calcStep = "Качество телеизмерений"
    EvaluateTelemetry wsRaw, g, res, pCol, timeCol, startRow, evEnd

    calcStep = "Качественный критерий"
    EvaluateQualitative wsRaw, st, g, res, pCol, fCol, timeCol

    calcStep = "Идентификация переходной функции h(15 с) (Прил. п. 6.2)"
    res.H15 = IdentifyH15(wsRaw, st, g, pCol, fCol, timeCol, firstExceedRow, p0, res.H15Note)

    calcStep = "Тип реакции (п. 4.4.1)"
    ClassifyReactionType wsRaw, st, g, res, pCol, fCol, timeCol, startRow, evEnd

    calcStep = "Итоговый вывод"
    res.Verdict = BuildGeneratorVerdict(g, res, st)
    If Left$(res.Verdict, 9) = "НАРУШЕНИЕ" Then AppendLog "WARN", g.Station & "/" & g.Generator, res.Verdict

    AnalyzeOneGenerator = res
    Exit Function

EH:
    Err.Raise vbObjectError + 2199, , "AnalyzeOneGenerator (" & g.Generator & " / " & calcStep & "): " & Err.Description
End Function

' ----------------------------------------------------------
' п. 4.2.6: среднее значение мощности на интервале [t(row) - windowSec ; t(row)].
' windowSec <= 0 -> текущее значение в строке row.
' ----------------------------------------------------------
Private Function AvgPowerBackward(ByVal wsRaw As Worksheet, ByVal pCol As Long, ByVal timeCol As Long, _
                                  ByVal row As Long, ByVal windowSec As Double) As Double
    Dim r1 As Long, r2 As Long, r As Long, sumP As Double, cnt As Long
    If windowSec <= 0 Then
        AvgPowerBackward = NzD(wsRaw.Cells(row, pCol).Value, 0)
        Exit Function
    End If
    WindowRowsAround wsRaw, timeCol, row, windowSec, r1, r2
    If r1 < 2 Then r1 = 2
    If r1 > row Then r1 = row
    For r = r1 To row
        If IsNumeric(wsRaw.Cells(r, pCol).Value) And Len(Trim$(CStr(wsRaw.Cells(r, pCol).Value))) > 0 Then
            sumP = sumP + CDbl(wsRaw.Cells(r, pCol).Value)
            cnt = cnt + 1
        End If
    Next r
    If cnt > 0 Then AvgPowerBackward = sumP / cnt Else AvgPowerBackward = NzD(wsRaw.Cells(row, pCol).Value, 0)
End Function

' ----------------------------------------------------------
' п. 4.2.11: наличие реакции регулятора частоты вращения турбины на
' первоначальное отклонение частоты - обратное изменению частоты изменение
' мощности через 15 с после выхода частоты за fнч. Расчёт по текущим значениям.
' Отсутствие реакции НА ЗАГРУЗКУ допустимо только при отсутствии запаса на
' открытие РК (резерв '+' по Pmax); реакция НА РАЗГРУЗКУ обязательна всегда.
' ----------------------------------------------------------
Private Sub EvaluateReaction15(ByVal wsRaw As Worksheet, ByRef st As TSettings, _
                               ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                               ByVal startRow As Long, ByVal firstExceedRow As Long, _
                               ByVal p0 As Double, ByVal pNom As Double, ByVal sPct As Double, _
                               ByVal fnch As Double, ByVal equipType As String, ByVal needSign As Integer, _
                               ByVal reservePlus As Double, ByVal reserveMinus As Double, _
                               ByRef status As String, ByRef pass As Boolean, _
                               ByRef dp15 As Double, ByRef note As String)
    Dim baseRow As Long, row15 As Long, dfr15 As Double, preq15 As Double, thrMW As Double
    dp15 = 0
    If needSign = 0 Then
        status = "Н/П"
        pass = True
        note = "Отклонение частоты в пределах fнч - оценка реакции РЧВ не требуется"
        Exit Sub
    End If
    baseRow = IIf(firstExceedRow > 0, firstExceedRow, startRow)
    row15 = RowByTimeOffset(wsRaw, timeCol, baseRow, 15#)
    dp15 = NzD(wsRaw.Cells(row15, pCol).Value, 0) - p0
    ' Требуемая мощность на 15-й секунде по текущей частоте (Kд(15 с))
    dfr15 = DeadbandDeviation(NzD(wsRaw.Cells(row15, fCol).Value, st.FNom) - st.FNom, fnch)
    preq15 = 0
    If dfr15 <> 0 And sPct > 0 Then
        preq15 = PReqCalc(sPct, pNom, st.FNom, DynamicKdByTime(equipType, 15#), dfr15, st.MaxReqPct)
        preq15 = ClampReserve(preq15, reservePlus, reserveMinus)
    End If
    thrMW = React15ThresholdMW(st, pNom, preq15)
    ClassifyReaction15 st, dp15, pNom, needSign, reservePlus, reserveMinus, thrMW, (dfr15 = 0), status, pass, note
End Sub

' Порог "наличия реакции": не более React15ThrPct %Pном и не более половины |Pтреб(15 с)|
' (чтобы при малом отклонении частоты не требовать больше, чем положено по Kд(15 с)).
Private Function React15ThresholdMW(ByRef st As TSettings, ByVal pNom As Double, ByVal preq15 As Double) As Double
    Dim thr As Double
    thr = st.React15ThrPct / 100# * pNom
    If thr < 0 Then thr = 0
    If Abs(preq15) > 0 And 0.5 * Abs(preq15) < thr Then thr = 0.5 * Abs(preq15)
    React15ThresholdMW = thr
End Function

' Классификация реакции по уже вычисленному dP15 (общая для генератора и суммы станции)
Private Sub ClassifyReaction15(ByRef st As TSettings, ByVal dp15 As Double, ByVal pNom As Double, _
                               ByVal needSign As Integer, ByVal reservePlus As Double, ByVal reserveMinus As Double, _
                               ByVal thrMW As Double, ByVal freqBackInside As Boolean, _
                               ByRef status As String, ByRef pass As Boolean, ByRef note As String)
    Dim dirText As String
    If needSign = 0 Then
        status = "Н/П"
        pass = True
        note = "Отклонение частоты в пределах fнч - оценка реакции РЧВ не требуется"
        Exit Sub
    End If
    dirText = IIf(needSign = 1, "загрузку", "разгрузку")

    If needSign * dp15 >= thrMW And Abs(dp15) > 0 Then
        status = "ОК"
        pass = True
        note = "Реакция РЧВ на " & dirText & " через 15 с есть: dP15 = " & Format(dp15, "0.000") & " МВт"
    ElseIf freqBackInside Then
        status = "Н/П"
        pass = True
        note = "К 15-й секунде частота вернулась в зону нечувствительности - реакция РЧВ не оценивается; dP15 = " & Format(dp15, "0.000") & " МВт"
    ElseIf needSign = 1 And reservePlus < thrMW Then
        status = "Н/П"
        pass = True
        note = "Нет запаса на открытие РК (резерв '+' = " & Format(reservePlus, "0.0") & _
               " МВт по Pmax) - отсутствие реакции на загрузку допустимо (п. 4.2.11); dP15 = " & Format(dp15, "0.000") & " МВт"
    Else
        status = "Нарушение"
        pass = False
        note = "Отсутствует реакция РЧВ на " & dirText & " через 15 с после выхода за fнч: dP15 = " & _
               Format(dp15, "0.000") & " МВт при пороге " & Format(needSign * thrMW, "0.000") & " МВт"
        If needSign = -1 And reserveMinus < thrMW Then
            note = note & " (резерв '-' = " & Format(reserveMinus, "0.0") & " МВт мал - проверьте Pmin; по п. 4.2.11 реакция на разгрузку обязательна)"
        End If
    End If
End Sub

' ----------------------------------------------------------
' п. 4.2.14.1 / Приложение (метод допустимых границ): дефицит первичной
' мощности по каждой точке окна, где частота вне fнч. Граница = Pтреб(t) по
' Kд(t), ограниченный резервом, минус допуск 1 %Pном (п. 4.2.7/4.2.8).
' Первые 15 с после старта (переходный участок) исключаются - они оцениваются
' по п. 4.2.11 (наличие реакции РЧВ).
' Возвращает макс. дефицит в %Pном и суммарную длительность дефицита, с.
' ----------------------------------------------------------
Private Sub EvaluateDeficit(ByVal wsRaw As Worksheet, ByRef st As TSettings, _
                            ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                            ByVal startRow As Long, ByVal endRow As Long, _
                            ByVal p0 As Double, ByVal pNom As Double, ByVal sPct As Double, _
                            ByVal fnch As Double, ByVal equipType As String, _
                            ByVal reservePlus As Double, ByVal reserveMinus As Double, _
                            ByRef defMaxPct As Double, ByRef defSec As Double)
    Dim r As Long, dFrT As Double, preqT As Double, allowed As Double, have As Double, deficit As Double
    Dim tSec As Double, dt As Double, tol As Double
    defMaxPct = 0
    defSec = 0
    If pNom <= 0 Or sPct <= 0 Then Exit Sub
    tol = 0.01 * pNom
    For r = startRow To endRow
        dFrT = DeadbandDeviation(NzD(wsRaw.Cells(r, fCol).Value, st.FNom) - st.FNom, fnch)
        If dFrT = 0 Then GoTo NextR
        tSec = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
        ' первые 15 с - переходный участок, оценивается отдельно по п. 4.2.11 (реакция РЧВ)
        If tSec < 15 Then GoTo NextR
        preqT = PReqCalc(sPct, pNom, st.FNom, DynamicKdByTime(equipType, tSec), dFrT, st.MaxReqPct)
        preqT = ClampReserve(preqT, reservePlus, reserveMinus)
        allowed = Abs(preqT) - tol
        If allowed <= 0 Then GoTo NextR
        have = SgnNZ(preqT) * (NzD(wsRaw.Cells(r, pCol).Value, 0) - p0)
        deficit = allowed - have
        If deficit > 0 Then
            If 100# * deficit / pNom > defMaxPct Then defMaxPct = 100# * deficit / pNom
            If r < endRow Then
                dt = SecBetween(wsRaw.Cells(r, timeCol).Value, wsRaw.Cells(r + 1, timeCol).Value)
            ElseIf r > startRow Then
                dt = SecBetween(wsRaw.Cells(r - 1, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
            Else
                dt = 1
            End If
            If dt <= 0 Or dt > 60 Then dt = 1
            defSec = defSec + dt
        End If
NextR:
    Next r
End Sub

' ----------------------------------------------------------
' Итоговый вывод по генератору (текст для Summary / листа графика)
' ----------------------------------------------------------
Private Function BuildGeneratorVerdict(ByRef g As TGenCfg, ByRef res As TGenResult, ByRef st As TSettings) As String
    Dim bad As String, remark As String, txt As String, okPart As String
    bad = ""
    remark = ""

    ' Событие не подлежит оценке по МУ
    If Not res.QuantApplicable And Not res.QualApplicable Then
        txt = "НЕ ОЦЕНИВАЕТСЯ: " & res.Applicability
        If Len(res.TelemetryNote) > 0 And Not res.TelemetryOK Then txt = txt & "; " & res.TelemetryNote
        BuildGeneratorVerdict = txt
        Exit Function
    End If

    ' --- Количественная оценка (п. 4.2) ---
    If res.QuantApplicable Then
        If res.Limited Then
            remark = remark & "Pтреб ограничен " & res.LimitType & " (п. 4.2.8: оценка по располагаемому резерву); "
        ElseIf Not res.QuantPass Then
            bad = bad & "Pфакт(tр) = " & Format(res.PFact, "0.00") & " МВт ниже Pтреб(tр) = " & Format(res.PReq, "0.00") & _
                  " МВт (" & Format(res.QuantPct, "0") & " %) - п. 4.2.14.1 (с учётом допуска п. 4.2.7/4.2.8 и Kд(t) Приложения); "
        End If
        If res.DeficitSec > 0 And res.QuantPass And Not res.Limited Then
            remark = remark & "кратковременный дефицит до " & Format(res.DeficitMaxPct, "0.0") & " %Pном в течение " & _
                     Format(res.DeficitSec, "0") & " с (п. 4.2.14.1, в пределах допуска в tр); "
        ElseIf res.DeficitSec > 0 Then
            remark = remark & "дефицит до " & Format(res.DeficitMaxPct, "0.0") & " %Pном в течение " & _
                     Format(res.DeficitSec, "0") & " с (п. 4.2.14.1); "
        End If
        If res.OscViolation Then
            bad = bad & res.OscNote & " - п. 4.2.14.4; "
        ElseIf Len(res.OscNote) > 0 And res.OscPeriods > 0 Then
            remark = remark & res.OscNote & "; "
        End If
    Else
        remark = remark & "количественная оценка не применяется (п. 4.1.2: " & res.Applicability & "); "
    End If

    ' --- Реакция РЧВ через 15 с (п. 4.2.11) ---
    If res.React15Status = "Нарушение" Then
        bad = bad & "отсутствует реакция РЧВ турбины через 15 с (dP15 = " & Format(res.DP15, "0.00") & " МВт) - п. 4.2.11; "
    ElseIf res.React15Status = "Н/П" And Len(res.React15Note) > 0 And res.Dfr <> 0 Then
        remark = remark & "реакция 15 с не оценивается (" & res.React15Note & "); "
    End If

    ' --- Качественная оценка (п. 4.4) ---
    If res.QualApplicable Then
        If Len(res.ReactionType) > 0 Then remark = remark & "тип реакции по п. 4.4.1: " & res.ReactionType & "; "
        If Left$(res.QualClauseVerdict, 9) = "НАРУШЕНИЕ" Then
            bad = bad & Mid$(res.QualClauseVerdict, 12) & "; "
        ElseIf Len(res.QualClauseVerdict) > 0 Then
            remark = remark & res.QualClauseVerdict & "; "
        End If
        If g.QualEnabled And res.Dfr <> 0 And Not res.QualPass Then
            remark = remark & "ступени динамики (Прил. 1, табл. Е): не выполнены " & res.QualFailedList & "; "
        End If
    End If

    If Len(res.TelemetryNote) > 0 And Not res.TelemetryOK Then remark = remark & res.TelemetryNote & "; "
    If res.H15 > -900 Then
        remark = remark & "h(15 с) = " & Format(res.H15, "0.00") & IIf(res.H15 < H15_REF, " < " & Format(H15_REF, "0.0") & " (ориентир)", "") & _
                 " (идентификация переходной функции, Прил. п. 6.2, справочно); "
    End If
    If g.ForcedNoRM Then
        remark = remark & "оборудование без РМ, режим «Авария»: принято S = " & Format(g.SPct, "0.0") & " %, fнч = " & _
                 Format(g.Fnch, "0.000") & " Гц (в Config: S = " & Format(g.SPctCfg, "0.0") & " %, fнч = " & Format(g.FnchCfg, "0.000") & " Гц); "
    End If

    If Len(bad) > 0 Then
        txt = "НАРУШЕНИЕ: " & bad
    Else
        okPart = "ОК: "
        If res.QuantApplicable Then okPart = okPart & "количественно (п. 4.2.14.1) и "
        okPart = okPart & "по реакции 15 с (п. 4.2.11) соответствует"
        If res.QualApplicable And res.ReactionCode = 1 Then
            okPart = okPart & "; качественно - реакция «адекватная» (п. 4.4.1)"
        ElseIf res.QualApplicable And res.ReactionCode > 1 Then
            okPart = okPart & "; качественно - тип реакции " & res.ReactionType & " (п. 4.4.1)"
        End If
        txt = okPart & "; "
    End If
    If Len(remark) > 0 Then txt = txt & "Замечания: " & remark
    If Right$(txt, 2) = "; " Then txt = Left$(txt, Len(txt) - 2)
    BuildGeneratorVerdict = txt
End Function

' ==========================================================
' 1.7.0: режим «Авария/Испытание», события, применимость, tр, колебания,
' телеизмерения, тип реакции, идентификация h(15 с), протокол
' ==========================================================

' Pтреб по формуле (2) МУ с ограничением |Pтреб| <= maxReqPct %Pном (Приложение: 0.1*Pном)
Private Function PReqCalc(ByVal sPct As Double, ByVal pNom As Double, ByVal fNom As Double, _
                          ByVal kd As Double, ByVal dfr As Double, ByVal maxReqPct As Double) As Double
    Dim v As Double, capMW As Double
    If dfr = 0 Or sPct <= 0 Or fNom <= 0 Then PReqCalc = 0: Exit Function
    v = -100# / sPct * pNom / fNom * kd * dfr
    If maxReqPct > 0 Then
        capMW = maxReqPct / 100# * pNom
        If v > capMW Then v = capMW
        If v < -capMW Then v = -capMW
    End If
    PReqCalc = v
End Function

Private Function ClampReserve(ByVal v As Double, ByVal reservePlus As Double, ByVal reserveMinus As Double) As Double
    If v > reservePlus Then v = reservePlus
    If v < -reserveMinus Then v = -reserveMinus
    ClampReserve = v
End Function

' Чтение строки Config с учётом режима: в «Аварии» для оборудования без РМ
' принимаются S = 6 %, fнч = 0.15 Гц независимо от Config.
Private Function ReadGenCfgEff(ByVal ws As Worksheet, ByVal r As Long, ByRef st As TSettings) As TGenCfg
    Dim g As TGenCfg
    g = ReadGenCfg(ws, r)
    g.SPctCfg = g.SPct
    g.FnchCfg = g.Fnch
    g.ForcedNoRM = False
    If st.IsEmergency And Not g.HasRM Then
        g.SPct = NO_RM_S_PCT
        g.Fnch = NO_RM_FNCH_HZ
        g.ForcedNoRM = True
    End If
    ReadGenCfgEff = g
End Function

' Среднее значение частоты на отрезке ±halfSec вокруг строки row (Приложение, п. 4.2.6)
' Границы окна ±halfSec вокруг строки row локальным сканированием (данные упорядочены по времени)
Private Sub WindowRowsAround(ByVal wsRaw As Worksheet, ByVal timeCol As Long, ByVal row As Long, _
                             ByVal halfSec As Double, ByRef r1 As Long, ByRef r2 As Long)
    Dim lastR As Long, t0 As Double, tv As Variant
    lastR = LastUsedRow(wsRaw)
    r1 = row: r2 = row
    If Not IsDate(wsRaw.Cells(row, timeCol).Value) Then Exit Sub
    t0 = CDbl(CDate(wsRaw.Cells(row, timeCol).Value)) * 86400#
    Do While r1 > 2
        tv = wsRaw.Cells(r1 - 1, timeCol).Value
        If Not IsDate(tv) Then Exit Do
        If t0 - CDbl(CDate(tv)) * 86400# > halfSec + 0.0001 Then Exit Do
        r1 = r1 - 1
    Loop
    Do While r2 < lastR
        tv = wsRaw.Cells(r2 + 1, timeCol).Value
        If Not IsDate(tv) Then Exit Do
        If CDbl(CDate(tv)) * 86400# - t0 > halfSec + 0.0001 Then Exit Do
        r2 = r2 + 1
    Loop
End Sub

Private Function MeanFreqAround(ByVal wsRaw As Worksheet, ByVal fCol As Long, ByVal timeCol As Long, _
                                ByVal row As Long, ByVal halfSec As Double, ByVal fNom As Double) As Double
    Dim r1 As Long, r2 As Long, r As Long, s As Double, n As Long
    WindowRowsAround wsRaw, timeCol, row, halfSec, r1, r2
    For r = r1 To r2
        If IsNumeric(wsRaw.Cells(r, fCol).Value) And Len(Trim$(CStr(wsRaw.Cells(r, fCol).Value))) > 0 Then
            s = s + CDbl(wsRaw.Cells(r, fCol).Value)
            n = n + 1
        End If
    Next r
    If n > 0 Then MeanFreqAround = s / n Else MeanFreqAround = fNom
End Function

' Квазиустановившийся режим по частоте: на отрезке ±15 с |f - fср| <= 10 мГц (Приложение)
Private Function IsQuasiSteady(ByVal wsRaw As Worksheet, ByVal fCol As Long, ByVal timeCol As Long, _
                               ByVal row As Long, ByVal fNom As Double, ByRef fMean As Double) As Boolean
    Dim r1 As Long, r2 As Long, r As Long, v As Variant
    fMean = MeanFreqAround(wsRaw, fCol, timeCol, row, QS_HALF_WINDOW_SEC, fNom)
    WindowRowsAround wsRaw, timeCol, row, QS_HALF_WINDOW_SEC, r1, r2
    For r = r1 To r2
        v = wsRaw.Cells(r, fCol).Value
        If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then
            If Abs(CDbl(v) - fMean) > QS_TOL_HZ + 0.0000001 Then IsQuasiSteady = False: Exit Function
        End If
    Next r
    IsQuasiSteady = True
End Function

' ----------------------------------------------------------
' Поиск событий: интервалов выхода частоты за зону нечувствительности fнч.
' Событие заканчивается, когда частота находится внутри fнч не менее
' SteadyWindowSec подряд, либо по истечении MaxEventSec. Короткие выбросы
' (< EVENT_MIN_OUTSIDE_SEC за fнч) не считаются событиями.
' При отключённом автопоиске - одно событие от заданного времени старта.
' Возвращает число событий (>= 1) и массив ev(1..n).
' ----------------------------------------------------------
Private Function FindEvents(ByVal wsRaw As Worksheet, ByVal timeCol As Long, ByVal fCol As Long, _
                            ByRef st As TSettings, ByVal fnch As Double, ByRef ev() As TEvent) As Long
    Dim lastR As Long, r As Long, n As Long
    Dim prevOut As Boolean, curOut As Boolean, inEvent As Boolean
    Dim lastOutRow As Long, tStart As Double, tCur As Double, tLastOut As Double
    Dim outsideSec As Double, bestRow As Long, bestAbs As Double, val As Double
    Dim dropped As Long

    lastR = LastUsedRow(wsRaw)
    ReDim ev(1 To 1)
    n = 0

    If lastR < 3 Then
        ev(1).StartRow = 2: ev(1).FirstExceedRow = 2: ev(1).EndRow = lastR
        FindEvents = 1
        Exit Function
    End If

    If Not st.AutoStart And st.EventStart > 0 Then
        bestRow = 2: bestAbs = 1E+99
        For r = 2 To lastR
            If IsDate(wsRaw.Cells(r, timeCol).Value) Then
                val = Abs(CDbl(CDate(wsRaw.Cells(r, timeCol).Value)) - st.EventStart)
                If val < bestAbs Then bestAbs = val: bestRow = r
            End If
        Next r
        ev(1).StartRow = bestRow
        ev(1).FirstExceedRow = bestRow
        ev(1).EndRow = RowByTimeOffset(wsRaw, timeCol, bestRow, st.MaxEventSec)
        FindEvents = 1
        Exit Function
    End If

    prevOut = (Abs(NzD(wsRaw.Cells(2, fCol).Value, st.FNom) - st.FNom) > fnch)
    inEvent = False
    For r = 3 To lastR
        curOut = (Abs(NzD(wsRaw.Cells(r, fCol).Value, st.FNom) - st.FNom) > fnch)
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then tCur = CDbl(CDate(wsRaw.Cells(r, timeCol).Value)) * 86400#
        If Not inEvent Then
            If curOut And Not prevOut Then
                inEvent = True
                n = n + 1
                ReDim Preserve ev(1 To n)
                ev(n).StartRow = r - 1
                ev(n).FirstExceedRow = r
                ev(n).EndRow = 0
                lastOutRow = r
                tStart = tCur
                tLastOut = tCur
                outsideSec = 0
            End If
        Else
            If curOut Then
                lastOutRow = r
                tLastOut = tCur
                outsideSec = outsideSec + 1
            End If
            ' конец события: частота внутри fнч дольше окна установившегося или достигнут предел окна
            If (Not curOut And (tCur - tLastOut) >= st.SteadyWindowSec) Or (tCur - tStart) >= st.MaxEventSec Then
                ev(n).EndRow = r
                inEvent = False
                ' фильтр коротких выбросов
                If SecBetween(wsRaw.Cells(ev(n).FirstExceedRow, timeCol).Value, wsRaw.Cells(lastOutRow, timeCol).Value) + 1 < EVENT_MIN_OUTSIDE_SEC Then
                    n = n - 1
                    dropped = dropped + 1
                    If n >= 1 Then ReDim Preserve ev(1 To n) Else ReDim ev(1 To 1)
                End If
            End If
        End If
        prevOut = curOut
    Next r
    If inEvent Then ev(n).EndRow = lastR

    If n = 0 Then
        ' Выхода за fнч не было: одно событие вокруг максимального отклонения (как в 1.6.x)
        bestRow = 2: bestAbs = -1
        For r = 2 To lastR
            val = NzD(wsRaw.Cells(r, fCol).Value, st.FNom) - st.FNom
            If Abs(val) > bestAbs Then bestAbs = Abs(val): bestRow = r
        Next r
        ReDim ev(1 To 1)
        ev(1).StartRow = bestRow
        ev(1).FirstExceedRow = bestRow
        ev(1).EndRow = RowByTimeOffset(wsRaw, timeCol, bestRow, st.MaxEventSec)
        n = 1
    End If

    If dropped > 0 Then AppendLog "INFO", "События", "Пропущено коротких выбросов частоты за fнч (< " & EVENT_MIN_OUTSIDE_SEC & " с): " & dropped
    If st.MaxEvents >= 1 And n > st.MaxEvents Then
        AppendLog "WARN", "События", "Найдено " & n & " событий; оцениваются первые " & st.MaxEvents & " (Config W15)"
        ReDim Preserve ev(1 To st.MaxEvents)
        n = st.MaxEvents
    End If
    FindEvents = n
End Function

' ----------------------------------------------------------
' Применимость оценки по п. 4.1.2 / 4.1.3 / 4.4.3 / 4.4.4 МУ.
' В режиме «Испытание» обе оценки считаются применимыми (контрольные испытания).
' ----------------------------------------------------------
Private Sub EvaluateApplicability(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, _
                                  ByVal fCol As Long, ByVal timeCol As Long, _
                                  ByVal startRow As Long, ByVal firstExceedRow As Long, ByVal evEnd As Long, _
                                  ByRef res As TGenResult)
    Dim r As Long, d As Double, dMax As Double, rMax As Long
    Dim run02 As Double, best02 As Double, outSec As Double, dt As Double
    Dim prevT As Double, curT As Double, minQuantSec As Double, isGpa As Boolean
    Dim txtQ As String, txtK As String

    dMax = 0: rMax = startRow
    run02 = 0: best02 = 0: outSec = 0
    prevT = -1
    For r = startRow To evEnd
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then
            curT = CDbl(CDate(wsRaw.Cells(r, timeCol).Value)) * 86400#
            If prevT >= 0 Then dt = curT - prevT Else dt = 1
            If dt <= 0 Or dt > 60 Then dt = 1
            prevT = curT
        Else
            dt = 1
        End If
        d = NzD(wsRaw.Cells(r, fCol).Value, st.FNom) - st.FNom
        If Abs(d) > Abs(dMax) Then dMax = d: rMax = r
        If Abs(d) >= QUANT_DF_HZ - 0.0000001 Then
            run02 = run02 + dt
            If run02 > best02 Then best02 = run02
        Else
            run02 = 0
        End If
        If Abs(d) > g.Fnch Then outSec = outSec + dt
    Next r

    res.DfMaxEvent = dMax
    res.DurAbove02Sec = best02
    res.DurOutsideSec = outSec
    If firstExceedRow > 0 And rMax >= firstExceedRow Then
        res.JumpSec = SecBetween(wsRaw.Cells(firstExceedRow, timeCol).Value, wsRaw.Cells(rMax, timeCol).Value)
    Else
        res.JumpSec = 0
    End If

    If Not st.IsEmergency Then
        res.QuantApplicable = True
        res.QualApplicable = True
        res.Applicability = "Испытание: колич. + кач. (применимость по п. 4.1.2/4.1.3 не проверяется)"
        Exit Sub
    End If

    isGpa = (InStr(UCase$(g.EquipType), "ГПА") > 0 Or InStr(UCase$(g.EquipType), "ДГУ") > 0)
    minQuantSec = IIf(isGpa, QUANT_MIN_SEC_GPA, QUANT_MIN_SEC_TES)

    ' п. 4.1.2: |dF| >= 0.2 Гц не менее 30 с (ТЭС) / 1 мин (ГПА, ДГУ)
    res.QuantApplicable = (best02 >= minQuantSec - 0.5)
    If res.QuantApplicable Then
        txtQ = "колич. (п. 4.1.2: |dF| >= 0.2 Гц в течение " & Format(best02, "0") & " с)"
    ElseIf best02 > 0 Then
        txtQ = "|dF| >= 0.2 Гц лишь " & Format(best02, "0") & " с (< " & Format(minQuantSec, "0") & " с по п. 4.1.2)"
    Else
        txtQ = "|dF|max = " & Format(Abs(dMax), "0.000") & " Гц < 0.2 Гц (п. 4.1.2)"
    End If

    ' п. 4.1.3: резкое (<= 10 с) отклонение продолжительностью >= 20 с;
    ' п. 4.4.4 (с РМ): выход за 50±0.15 Гц; п. 4.4.3 (без РМ): скачок ±0.2..0.4 Гц.
    Dim jumpOk As Boolean, durOk As Boolean, levelOk As Boolean
    jumpOk = (res.JumpSec <= QUAL_JUMP_SEC + 0.5) Or res.QuantApplicable
    durOk = (outSec >= QUAL_MIN_SEC - 0.5)
    If g.HasRM Then
        levelOk = (Abs(dMax) > QUAL_RM_DF_HZ)
    Else
        levelOk = (Abs(dMax) >= QUANT_DF_HZ - 0.0000001 And Abs(dMax) <= QUAL_NORM_DF_MAX_HZ + 0.0000001)
    End If
    res.QualApplicable = jumpOk And durOk And levelOk
    If res.QualApplicable Then
        txtK = "кач. (п. 4.1.3, " & IIf(g.HasRM, "п. 4.4.4: с РМ, |dF| > 0.15 Гц", "п. 4.4.3: без РМ, |dF| в 0.2..0.4 Гц") & _
               "; нарастание " & Format(res.JumpSec, "0") & " с, вне fнч " & Format(outSec, "0") & " с)"
    Else
        txtK = "кач. не применяется ("
        If Not jumpOk Then txtK = txtK & "нарастание " & Format(res.JumpSec, "0") & " с > 10 с; "
        If Not durOk Then txtK = txtK & "вне fнч лишь " & Format(outSec, "0") & " с < 20 с; "
        If Not levelOk Then txtK = txtK & "|dF|max = " & Format(Abs(dMax), "0.000") & " Гц вне диапазона " & _
            IIf(g.HasRM, "> 0.15 Гц (п. 4.4.4)", "0.2..0.4 Гц (п. 4.4.3)") & "; "
        If Right$(txtK, 2) = "; " Then txtK = Left$(txtK, Len(txtK) - 2)
        txtK = txtK & ")"
    End If

    If res.QuantApplicable And res.QualApplicable Then
        res.Applicability = "Колич. + кач.: " & txtQ & "; " & txtK
    ElseIf res.QuantApplicable Then
        res.Applicability = "Только колич.: " & txtQ & "; " & txtK
    ElseIf res.QualApplicable Then
        res.Applicability = "Только кач.: " & txtQ & "; " & txtK
    Else
        res.Applicability = "отклонение частоты не удовлетворяет п. 4.1.2/4.1.3 МУ: " & txtQ & "; " & txtK
    End If
End Sub

' ----------------------------------------------------------
' Выбор момента оценки tр (п. 4.2.5, 4.2.9, Приложение):
' среди квазиустановившихся по частоте точек окна события (t >= 15 с от t0,
' |fср - fном| > fнч) выбирается точка с максимальным дефицитом первичной
' мощности относительно границы Pтреб(t) - 1 %Pном. Если дефицита нет -
' точка с максимальным |dFр|. Если квазиустановившихся точек нет - конец
' контрольного интервала (Config W5), как в 1.6.x.
' ----------------------------------------------------------
Private Sub SelectEvaluationPoint(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, _
                                  ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                                  ByVal startRow As Long, ByVal firstExceedRow As Long, ByVal evEnd As Long, _
                                  ByVal endRef As Long, ByVal p0 As Double, _
                                  ByVal reservePlus As Double, ByVal reserveMinus As Double, ByRef res As TGenResult)
    Dim r As Long, fMean As Double, dfrT As Double, tSec As Double, preqT As Double
    Dim have As Double, deficit As Double, tol As Double
    Dim bestRow As Long, bestDef As Double, bestDfrRow As Long, bestDfr As Double
    Dim anyQS As Boolean, stepR As Long, lastR As Long

    tol = 0.01 * g.PNom
    bestRow = 0: bestDef = -1E+99
    bestDfrRow = 0: bestDfr = 0
    anyQS = False
    lastR = LastUsedRow(wsRaw)
    If evEnd > lastR Then evEnd = lastR

    ' шаг перебора - по строкам; для длинных событий ограничиваем число проверок ~600
    stepR = 1
    If evEnd - startRow > 600 Then stepR = (evEnd - startRow) \ 600 + 1

    For r = startRow To evEnd Step stepR
        tSec = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
        If tSec < 15 Then GoTo NextR
        If Not IsQuasiSteady(wsRaw, fCol, timeCol, r, st.FNom, fMean) Then GoTo NextR
        dfrT = DeadbandDeviation(fMean - st.FNom, g.Fnch)
        If dfrT = 0 Then GoTo NextR
        anyQS = True
        preqT = ClampReserve(PReqCalc(g.SPct, g.PNom, st.FNom, DynamicKdByTime(g.EquipType, tSec), dfrT, st.MaxReqPct), reservePlus, reserveMinus)
        have = SgnNZ(preqT) * (AvgPowerBackward(wsRaw, pCol, timeCol, r, st.AvgWindowSec) - p0)
        deficit = (Abs(preqT) - tol) - have
        If deficit > bestDef Then bestDef = deficit: bestRow = r
        If Abs(dfrT) > Abs(bestDfr) Then bestDfr = dfrT: bestDfrRow = r
NextR:
    Next r

    If anyQS Then
        If bestDef > 0 Then
            res.TrRow = bestRow
        Else
            ' дефицита нигде нет (п. 4.2.5: момент может выбираться любым) - берём макс. |dFр|
            res.TrRow = IIf(bestDfrRow > 0, bestDfrRow, bestRow)
        End If
        res.TrQuasiSteady = True
    Else
        res.TrRow = endRef
        res.TrQuasiSteady = False
    End If
    If res.TrRow < startRow Then res.TrRow = startRow
    res.TrTime = wsRaw.Cells(res.TrRow, timeCol).Value
    res.TrSec = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(res.TrRow, timeCol).Value)
End Sub

' ----------------------------------------------------------
' п. 4.2.6 / 4.2.14.4: незатухающие (слабозатухающие) колебания мощности -
' более 5 подряд периодов с амплитудой > 1 %Pном, не связанные с колебаниями
' частоты. Экстремумы ищутся по сглаженному (5 точек) ряду dP; колебание
' считается связанным с частотой, если у большинства экстремумов мощности
' в пределах ±3 с есть экстремум частоты с размахом > 10 мГц.
' ----------------------------------------------------------
Private Sub EvaluateOscillations(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
                                 ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                                 ByVal startRow As Long, ByVal evEnd As Long)
    Dim n As Long, i As Long, r As Long
    Dim p() As Double, f() As Double, ps() As Double
    Dim extrIdx() As Long, extrCnt As Long
    Dim prevDir As Integer, curDir As Integer, dP As Double, noise As Double
    Dim amp As Double, maxAmp As Double, cntBig As Long, bestRun As Long, run As Long
    Dim fTied As Long, j As Long, fLo As Double, fHi As Double, k As Long
    Dim s As Double, c As Long

    res.OscViolation = False
    res.OscPeriods = 0
    res.OscAmplPct = 0
    res.OscNote = ""
    n = evEnd - startRow + 1
    If n < 20 Or g.PNom <= 0 Then Exit Sub
    If n > 5000 Then n = 5000

    ReDim p(1 To n): ReDim f(1 To n): ReDim ps(1 To n)
    For i = 1 To n
        r = startRow + i - 1
        p(i) = NzD(wsRaw.Cells(r, pCol).Value, 0)
        f(i) = NzD(wsRaw.Cells(r, fCol).Value, st.FNom)
    Next i
    ' сглаживание 5 точек
    For i = 1 To n
        s = 0: c = 0
        For j = i - 2 To i + 2
            If j >= 1 And j <= n Then s = s + p(j): c = c + 1
        Next j
        ps(i) = s / c
    Next i

    noise = 0.002 * g.PNom   ' 0.2 %Pном - порог значимого изменения
    ReDim extrIdx(1 To n)
    extrCnt = 0
    prevDir = 0
    For i = 2 To n
        dP = ps(i) - ps(i - 1)
        If dP > noise Then
            curDir = 1
        ElseIf dP < -noise Then
            curDir = -1
        Else
            curDir = 0
        End If
        If prevDir <> 0 And curDir <> 0 And curDir <> prevDir Then
            extrCnt = extrCnt + 1
            extrIdx(extrCnt) = i - 1
        End If
        If curDir <> 0 Then prevDir = curDir
    Next i
    If extrCnt < 2 Then Exit Sub

    ' амплитуда полупериодов (|разность соседних экстремумов| / 2); подряд идущие "большие" полуволны
    maxAmp = 0: cntBig = 0: bestRun = 0: run = 0
    For i = 2 To extrCnt
        amp = Abs(ps(extrIdx(i)) - ps(extrIdx(i - 1))) / 2#
        If amp > maxAmp Then maxAmp = amp
        If amp > OSC_MIN_AMPL_PCT / 100# * g.PNom Then
            run = run + 1
            If run > bestRun Then bestRun = run
            cntBig = cntBig + 1
        Else
            run = 0
        End If
    Next i
    res.OscAmplPct = 100# * maxAmp / g.PNom
    res.OscPeriods = bestRun \ 2    ' два полупериода = один период

    ' связь с частотой: вблизи экстремума мощности (±3 с) есть размах частоты > 10 мГц
    fTied = 0
    For i = 1 To extrCnt
        fLo = 1E+99: fHi = -1E+99
        r = startRow + extrIdx(i) - 1
        k = extrIdx(i)
        Do While k >= 1
            If Abs(SecBetween(wsRaw.Cells(r, timeCol).Value, wsRaw.Cells(startRow + k - 1, timeCol).Value)) > 3.0001 Then Exit Do
            If f(k) < fLo Then fLo = f(k)
            If f(k) > fHi Then fHi = f(k)
            k = k - 1
        Loop
        k = extrIdx(i) + 1
        Do While k <= n
            If Abs(SecBetween(wsRaw.Cells(r, timeCol).Value, wsRaw.Cells(startRow + k - 1, timeCol).Value)) > 3.0001 Then Exit Do
            If f(k) < fLo Then fLo = f(k)
            If f(k) > fHi Then fHi = f(k)
            k = k + 1
        Loop
        If fHi - fLo > QS_TOL_HZ Then fTied = fTied + 1
    Next i

    If res.OscPeriods >= OSC_MIN_PERIODS And res.OscAmplPct > OSC_MIN_AMPL_PCT Then
        If fTied * 2 > extrCnt Then
            res.OscNote = "колебания мощности (" & res.OscPeriods & " периодов, амплитуда до " & Format(res.OscAmplPct, "0.0") & _
                          " %Pном) связаны с колебаниями частоты - не нарушение по п. 4.2.14.4"
        Else
            res.OscViolation = True
            res.OscNote = "незатухающие колебания мощности: " & res.OscPeriods & " периодов подряд с амплитудой до " & _
                          Format(res.OscAmplPct, "0.0") & " %Pном (> 1 %Pном), не связанные с частотой"
        End If
    ElseIf res.OscPeriods > 0 And res.OscAmplPct > OSC_MIN_AMPL_PCT Then
        res.OscNote = "колебания мощности " & res.OscPeriods & " пер., амплитуда до " & Format(res.OscAmplPct, "0.0") & " %Pном (п. 4.2.6, < 5 периодов)"
    End If
End Sub

' ----------------------------------------------------------
' Качество телеизмерений мощности в окне события: пропуски, "замороженные"
' значения (более 11 одинаковых подряд - по аналогии с п. 6.1.8.3 Приложения 1),
' одиночные скачки > 20 %Pном за один шаг.
' ----------------------------------------------------------
Private Sub EvaluateTelemetry(ByVal wsRaw As Worksheet, ByRef g As TGenCfg, ByRef res As TGenResult, _
                              ByVal pCol As Long, ByVal timeCol As Long, ByVal startRow As Long, ByVal evEnd As Long)
    Dim r As Long, r1 As Long, total As Long, blanks As Long, run As Long, maxRun As Long
    Dim prevV As Double, curV As Double, hasPrev As Boolean, spikes As Long, note As String
    Dim lastR As Long
    res.TelemetryOK = True
    res.TelemetryNote = ""
    lastR = LastUsedRow(wsRaw)
    r1 = RowByTimeOffset(wsRaw, timeCol, startRow, -15#)
    If r1 < 2 Then r1 = 2
    If evEnd > lastR Then evEnd = lastR
    hasPrev = False
    For r = r1 To evEnd
        total = total + 1
        If Not IsNumeric(wsRaw.Cells(r, pCol).Value) Or Len(Trim$(CStr(wsRaw.Cells(r, pCol).Value))) = 0 Then
            blanks = blanks + 1
            run = 0
            hasPrev = False
        Else
            curV = CDbl(wsRaw.Cells(r, pCol).Value)
            If hasPrev Then
                If Abs(curV - prevV) < 0.0000001 Then
                    run = run + 1
                    If run > maxRun Then maxRun = run
                Else
                    run = 0
                End If
                If g.PNom > 0 Then
                    If Abs(curV - prevV) > 0.2 * g.PNom Then spikes = spikes + 1
                End If
            End If
            prevV = curV
            hasPrev = True
        End If
    Next r
    If total = 0 Then Exit Sub
    note = ""
    If blanks > 0 Then
        note = note & "пропуски ТИ мощности: " & blanks & " из " & total & " точек; "
        If blanks * 10 > total Then res.TelemetryOK = False
    End If
    If maxRun + 1 > TELEM_MAX_REPEAT Then
        note = note & "одинаковые значения мощности " & (maxRun + 1) & " точек подряд (> " & TELEM_MAX_REPEAT & ", возможное «замораживание» ТИ); "
        res.TelemetryOK = False
    End If
    If spikes > 0 Then
        note = note & "скачки мощности > 20 %Pном за шаг: " & spikes & "; "
        If spikes > 2 Then res.TelemetryOK = False
    End If
    If Len(note) > 0 Then
        If Right$(note, 2) = "; " Then note = Left$(note, Len(note) - 2)
        res.TelemetryNote = IIf(res.TelemetryOK, "ТИ: ", "ТИ сомнительны: ") & note
    End If
End Sub

' ----------------------------------------------------------
' Тип реакции по п. 4.4.1 МУ (1 «адекватная» ... 8 «телеизмерение») и вывод по
' п. 4.4.3 (без РМ: фиксируется факт отсутствия «адекватной» реакции; нарушение -
' при трёх и более фактах за год) / п. 4.4.4 (с РМ: нарушение при отсутствии
' «адекватной» реакции при наличии резерва).
' ----------------------------------------------------------
Private Sub ClassifyReactionType(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
                                 ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                                 ByVal startRow As Long, ByVal evEnd As Long)
    Dim needSign As Integer, reserveNeed As Double, thr1 As Double
    Dim r As Long, dP As Double, maxGood As Double, maxBad As Double, rMaxGood As Long
    Dim dPTr As Double, code As Long, tName As String
    Dim lastR As Long

    res.ReactionCode = 0
    res.ReactionType = ""
    res.QualClauseVerdict = ""
    needSign = SgnNZ(res.PReqOrig)
    If needSign = 0 Then needSign = -SgnNZ(res.DfMaxEvent)
    If needSign = 0 Then Exit Sub

    thr1 = 0.01 * g.PNom
    reserveNeed = IIf(needSign = 1, res.ReservePlus, res.ReserveMinus)
    lastR = LastUsedRow(wsRaw)
    If evEnd > lastR Then evEnd = lastR

    maxGood = 0: maxBad = 0: rMaxGood = startRow
    For r = startRow To evEnd
        dP = needSign * (NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0)
        If dP > maxGood Then maxGood = dP: rMaxGood = r
        If -dP > maxBad Then maxBad = -dP
    Next r
    dPTr = needSign * res.PFact

    If Not res.TelemetryOK And maxGood < thr1 And maxBad < thr1 Then
        code = 8: tName = "телеизмерение"
    ElseIf reserveNeed < thr1 Then
        code = 7: tName = "без резерва"
    ElseIf res.OscViolation Then
        code = 5: tName = "колебания"
    ElseIf maxBad > thr1 And maxGood < thr1 Then
        code = 4: tName = "противоположная"
    ElseIf maxGood < thr1 Then
        code = 6: tName = "нет реакции"
    ElseIf res.React15Status = "ОК" And dPTr < 0.5 * maxGood And dPTr < Abs(res.PReq) - thr1 And res.Dfr <> 0 Then
        code = 2: tName = "с провалом"
    ElseIf res.React15Status = "Нарушение" And dPTr >= 0.5 * Abs(res.PReqStatic) Then
        code = 3: tName = "котельная"
    ElseIf res.React15Status = "Нарушение" And maxGood >= thr1 Then
        code = 3: tName = "котельная"
    Else
        code = 1: tName = "адекватная"
    End If
    res.ReactionCode = code
    res.ReactionType = code & " - «" & tName & "»"

    If Not res.QualApplicable Then Exit Sub
    If code = 1 Then
        res.QualClauseVerdict = "реакция «адекватная» - требования п. " & IIf(g.HasRM, "4.4.4", "4.4.3") & " выполнены"
    ElseIf code = 7 Then
        res.QualClauseVerdict = "резерв на изменение мощности отсутствует - отсутствие реакции не является нарушением (п. " & IIf(g.HasRM, "4.4.4", "4.4.3") & ")"
    ElseIf code = 8 Then
        res.QualClauseVerdict = "качественная оценка невозможна из-за качества ТИ (тип 8 по п. 4.4.1)"
    ElseIf g.HasRM Then
        res.QualClauseVerdict = "НАРУШЕНИЕ: отсутствие реакции типа «адекватная» (тип " & res.ReactionType & ") при наличии резерва " & _
                                Format(reserveNeed, "0.0") & " МВт - п. 4.4.4"
    Else
        res.QualClauseVerdict = "зафиксирован факт отсутствия реакции типа «адекватная» (тип " & res.ReactionType & ") при резерве " & _
                                Format(reserveNeed, "0.0") & " МВт - п. 4.4.3 (нарушение фиксируется при >= 3 фактах за календарный год)"
    End If
End Sub

' ----------------------------------------------------------
' Приложение к МУ, п. 6.2: идентификация переходной функции генерирующего
' оборудования по отрезку [t-15; t+30] с (46 точек с шагом 1 с) относительно
' момента выхода частоты за fнч. x = Pтреб(Kд=1)/Pном*100 [%], y = dP/Pном*100 [%],
' сглаживание скользящим средним по 5 точкам, приращения dx/dy, СЛАУ 15x15
' (нормальные уравнения) -> приращения dh(1..15), h(15) = сумма dh.
' Возвращает h(15 с) или -999, если вычислить невозможно. Результат справочный.
' ----------------------------------------------------------
Private Function IdentifyH15(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, _
                             ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long, _
                             ByVal centerRow As Long, ByVal p0 As Double, ByRef note As String) As Double
    Const NPTS As Long = 50      ' t-17 .. t+32 (46 точек + по 2 для сглаживания)
    Const NH As Long = 15
    Dim x() As Double, y() As Double, xs() As Double, ys() As Double, dx() As Double, dy() As Double
    Dim a() As Double, b() As Double, dh() As Double
    Dim i As Long, j As Long, k As Long, r As Long, lastR As Long
    Dim dfr As Double, s As Double, c As Long, sumAbsDx As Double, h As Double, tr As Double
    ReDim x(1 To NPTS): ReDim y(1 To NPTS): ReDim xs(1 To NPTS): ReDim ys(1 To NPTS)
    ReDim dx(1 To NPTS): ReDim dy(1 To NPTS)
    ReDim a(1 To NH, 1 To NH): ReDim b(1 To NH): ReDim dh(1 To NH)

    IdentifyH15 = -999
    note = ""
    If centerRow <= 0 Or g.PNom <= 0 Or g.SPct <= 0 Then note = "h(15 с): нет момента выхода за fнч": Exit Function
    lastR = LastUsedRow(wsRaw)
    If RowByTimeOffset(wsRaw, timeCol, centerRow, 32#) >= lastR Then note = "h(15 с): недостаточно данных после выхода за fнч": Exit Function

    For i = 1 To NPTS
        r = RowByTimeOffset(wsRaw, timeCol, centerRow, CDbl(i - 18))   ' i=18 -> t
        If r < 2 Then r = 2
        dfr = DeadbandDeviation(NzD(wsRaw.Cells(r, fCol).Value, st.FNom) - st.FNom, g.Fnch)
        x(i) = PReqCalc(g.SPct, g.PNom, st.FNom, 1#, dfr, st.MaxReqPct) / g.PNom * 100#
        y(i) = (NzD(wsRaw.Cells(r, pCol).Value, p0) - p0) / g.PNom * 100#
    Next i
    ' скользящее среднее по 5 точкам
    For i = 1 To NPTS
        s = 0: c = 0
        For j = i - 2 To i + 2
            If j >= 1 And j <= NPTS Then s = s + x(j): c = c + 1
        Next j
        xs(i) = s / c
        s = 0: c = 0
        For j = i - 2 To i + 2
            If j >= 1 And j <= NPTS Then s = s + y(j): c = c + 1
        Next j
        ys(i) = s / c
    Next i
    ' приращения на рабочем отрезке i = 3..48 (46 точек)
    dx(3) = 0: dy(3) = 0
    sumAbsDx = 0
    For i = 4 To 48
        dx(i) = xs(i) - xs(i - 1)
        dy(i) = ys(i) - ys(i - 1)
        sumAbsDx = sumAbsDx + Abs(dx(i))
    Next i
    If sumAbsDx < 0.05 Then note = "h(15 с): входное воздействие (Pтреб) практически не менялось - идентификация не выполнена": Exit Function

    ' нормальные уравнения: A(i,j) = sum_k dx(k-i)*dx(k-j), b(i) = sum_k dx(k-i)*dy(k)
    For i = 1 To NH
        b(i) = 0
        For j = 1 To NH
            a(i, j) = 0
        Next j
    Next i
    For k = 3 + NH To 48
        For i = 1 To NH
            b(i) = b(i) + dx(k - i) * dy(k)
            For j = 1 To NH
                a(i, j) = a(i, j) + dx(k - i) * dx(k - j)
            Next j
        Next i
    Next k
    ' регуляризация (малая добавка к диагонали) для устойчивости при ступенчатом входе
    tr = 0
    For i = 1 To NH
        tr = tr + a(i, i)
    Next i
    For i = 1 To NH
        a(i, i) = a(i, i) + 0.000001 * (tr / NH + 0.000001)
    Next i
    If Not SolveLinearSystem(a, b, dh, NH) Then note = "h(15 с): СЛАУ вырождена": Exit Function
    h = 0
    For i = 1 To NH
        h = h + dh(i)
    Next i
    IdentifyH15 = h
    note = "h(15 с) = " & Format(h, "0.00") & " (Прил. п. 6.2, справочно; ориентир " & Format(H15_REF, "0.0") & ")"
End Function

' Решение СЛАУ методом Гаусса с выбором главного элемента. a/b модифицируются.
Private Function SolveLinearSystem(ByRef a() As Double, ByRef b() As Double, ByRef x() As Double, ByVal n As Long) As Boolean
    Dim i As Long, j As Long, k As Long, piv As Long, tmp As Double, factor As Double
    For k = 1 To n
        piv = k
        For i = k + 1 To n
            If Abs(a(i, k)) > Abs(a(piv, k)) Then piv = i
        Next i
        If Abs(a(piv, k)) < 1E-14 Then SolveLinearSystem = False: Exit Function
        If piv <> k Then
            For j = 1 To n
                tmp = a(k, j): a(k, j) = a(piv, j): a(piv, j) = tmp
            Next j
            tmp = b(k): b(k) = b(piv): b(piv) = tmp
        End If
        For i = k + 1 To n
            factor = a(i, k) / a(k, k)
            For j = k To n
                a(i, j) = a(i, j) - factor * a(k, j)
            Next j
            b(i) = b(i) - factor * b(k)
        Next i
    Next k
    For i = n To 1 Step -1
        tmp = b(i)
        For j = i + 1 To n
            tmp = tmp - a(i, j) * x(j)
        Next j
        x(i) = tmp / a(i, i)
    Next i
    SolveLinearSystem = True
End Function

' ----------------------------------------------------------
' Протокол (текстовый отчёт)
' ----------------------------------------------------------
Private Sub AddProtocolLine(ByVal s As String)
    If m_Protocol Is Nothing Then Set m_Protocol = New Collection
    m_Protocol.Add s
End Sub

Private Sub AddProtocolGenerator(ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim hdr As String
    hdr = g.Station & " / " & g.Generator & " (" & g.EquipType & ", Pном " & Format(g.PNom, "0.0") & " МВт, S " & Format(g.SPct, "0.0") & _
          " %, fнч " & Format(g.Fnch, "0.000") & " Гц, " & IIf(g.HasRM, "с РМ", "без РМ") & ", " & g.Layout & ")"
    If res.EventIdx > 1 Then hdr = hdr & " - событие " & res.EventIdx
    AddProtocolLine hdr
    AddProtocolLine "  Старт t0: " & Format(res.StartTime, "dd.mm.yyyy hh:mm:ss") & "; выход за fнч: " & _
        IIf(IsDate(res.FirstExceedTime), Format(res.FirstExceedTime, "hh:mm:ss"), "-") & "; tр: " & _
        IIf(IsDate(res.TrTime), Format(res.TrTime, "hh:mm:ss"), "-") & " (" & Format(res.TrSec, "0") & " с, " & _
        IIf(res.TrQuasiSteady, "квазиустановившийся режим", "переходный режим") & ")"
    AddProtocolLine "  Применимость: " & res.Applicability
    AddProtocolLine "  dF(tр) = " & Format(res.Dfr + IIf(res.Dfr > 0, g.Fnch, IIf(res.Dfr < 0, -g.Fnch, 0)), "0.000") & " Гц; dFр = " & Format(res.Dfr, "0.000") & _
        " Гц; Kд(tр) = " & Format(res.KdUsedQuant, "0.00") & "; P0 = " & Format(res.P0, "0.000") & " МВт; Pтек(tр) = " & Format(res.PTek, "0.000") & _
        " МВт; Pтреб = " & Format(res.PReq, "0.000") & " МВт" & IIf(res.Limited, " (ограничен " & res.LimitType & ")", "") & _
        "; Pфакт = " & Format(res.PFact, "0.000") & " МВт; выполнение " & Format(res.QuantPct, "0") & " %"
    AddProtocolLine "  Реакция РЧВ 15 с (п. 4.2.11): " & res.React15Status & " (dP15 = " & Format(res.DP15, "0.000") & " МВт)" & _
        "; дефицит (п. 4.2.14.1): " & IIf(res.DeficitSec > 0, "до " & Format(res.DeficitMaxPct, "0.0") & " %Pном, " & Format(res.DeficitSec, "0") & " с", "нет") & _
        "; тип реакции (п. 4.4.1): " & IIf(Len(res.ReactionType) > 0, res.ReactionType, "-")
    AddProtocolLine "  ВЫВОД: " & res.Verdict
    AddProtocolLine ""
End Sub

Private Sub AddProtocolStation(ByRef sr As TStationResult)
    AddProtocolLine sr.Station & IIf(Len(sr.Paroprovod) > 0, " / паропровод " & sr.Paroprovod, "") & " - СУММА (" & sr.Count & " ген., " & _
        IIf(sr.IsNeblok, "неблочная часть, п. 4.2.10", "блочная компоновка - справочно") & ")" & IIf(sr.EventIdx > 1, " - событие " & sr.EventIdx, "")
    AddProtocolLine "  Старт t0: " & Format(sr.StartTime, "dd.mm.yyyy hh:mm:ss") & "; tр: " & IIf(IsDate(sr.TrTime), Format(sr.TrTime, "hh:mm:ss"), "-") & _
        " (" & Format(sr.TrSec, "0") & " с); P0 = " & Format(sr.P0, "0.000") & " МВт; Pтреб_сум = " & Format(sr.PReq, "0.000") & _
        " МВт; Pфакт_сум = " & Format(sr.PFact, "0.000") & " МВт; выполнение " & Format(sr.QuantPct, "0") & " %"
    AddProtocolLine "  Применимость: " & sr.Applicability
    AddProtocolLine "  Реакция РЧВ 15 с: " & sr.React15Status & " (dP15 = " & Format(sr.DP15, "0.000") & " МВт); дефицит: " & _
        IIf(sr.DeficitSec > 0, "до " & Format(sr.DeficitMaxPct, "0.0") & " %Pном_сум, " & Format(sr.DeficitSec, "0") & " с", "нет")
    AddProtocolLine "  ВЫВОД: " & sr.Verdict
    AddProtocolLine ""
End Sub

Private Sub WriteProtocolSheet(ByRef st As TSettings, ByVal wsRaw As Worksheet)
    Dim ws As Worksheet, r As Long, i As Long, lastR As Long
    Dim tFirst As Variant, tLast As Variant
    Set ws = EnsureSheet(SH_PROT)
    ws.Cells.Clear
    lastR = LastUsedRow(wsRaw)
    tFirst = wsRaw.Cells(2, 1).Value
    tLast = wsRaw.Cells(lastR, 1).Value
    ws.Range("A1").Value = "ПРОТОКОЛ оценки участия генерирующего оборудования в ОПРЧ"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "Методика: Методические указания по оценке участия генерирующего оборудования в ОПРЧ (Приложение 14 к Техтребованиям); " & _
                           "ступени динамики - Приложение 1 (МУ проверки готовности)."
    ws.Range("A3").Value = "Программа: OPRCH_Monitor v" & OPRCH_VERSION & "; режим: " & st.Mode & "; дата расчёта: " & Format(Now, "dd.mm.yyyy hh:mm")
    ws.Range("A4").Value = "Выборка RawData: " & IIf(IsDate(tFirst), Format(tFirst, "dd.mm.yyyy hh:mm:ss"), "-") & " - " & _
                           IIf(IsDate(tLast), Format(tLast, "dd.mm.yyyy hh:mm:ss"), "-") & " (" & (lastR - 1) & " строк); fном = " & st.FNom & " Гц; " & _
                           "контрольный интервал " & st.QuantIntervalSec & " с; допуск колич. " & st.QuantTolPct & " %; ограничение Pтреб " & st.MaxReqPct & " %Pном"
    ws.Range("A5").Value = IIf(st.IsEmergency, _
        "Режим «Авария»: для оборудования без РМ приняты S = " & NO_RM_S_PCT & " %, fнч = " & NO_RM_FNCH_HZ & " Гц; применимость оценок - по п. 4.1.2/4.1.3 МУ.", _
        "Режим «Испытание»: использованы фактические S/fнч из Config; применимость по п. 4.1.2/4.1.3 не проверялась.")
    ws.Range("A6").Value = "Нарушения формулируются со ссылкой на пункты МУ: 4.2.11 (реакция РЧВ 15 с), 4.2.14.1 (дефицит первичной мощности), " & _
                           "4.2.14.4 (колебания), 4.4.3/4.4.4 (отсутствие «адекватной» реакции без РМ / с РМ), 4.2.10 (оценка неблочной части по сумме)."
    r = 8
    If Not m_Protocol Is Nothing Then
        For i = 1 To m_Protocol.Count
            ws.Cells(r, 1).Value = CStr(m_Protocol(i))
            If Left$(CStr(m_Protocol(i)), 2) <> "  " And Len(CStr(m_Protocol(i))) > 0 Then ws.Cells(r, 1).Font.Bold = True
            If InStr(CStr(m_Protocol(i)), "ВЫВОД: НАРУШЕНИЕ") > 0 Then ws.Cells(r, 1).Font.Color = RGB(192, 0, 0)
            If InStr(CStr(m_Protocol(i)), "ВЫВОД: ОК") > 0 Then ws.Cells(r, 1).Font.Color = RGB(0, 97, 0)
            r = r + 1
        Next i
    End If
    ws.Cells(r + 1, 1).Value = "Не автоматизировано (требует ручной проверки): срабатывание технологических защит / вмешательство персонала (п. 4.2.14.1, абз. 2), " & _
                               "влияние команд диспетчера и АРЧМ (п. 4.2.9), правило трёх фактов за год для оборудования без РМ (п. 4.4.3), " & _
                               "заявки на вывод из ОПРЧ / пуски и остановы (п. 4.1.6)."
    ws.Columns("A").ColumnWidth = 160
    ws.Range("A1:A" & (r + 1)).WrapText = True
    ws.Range("A1:A" & (r + 1)).VerticalAlignment = xlTop
End Sub

' Нестрогий поиск колонки RawData: сравнение после нормализации (регистр, пробелы,
' '_'/'-'/'.', латиница P/C/T/B/H/K/M/O/A/E/X -> кириллица); при неудаче - вхождение.
Private Function FindHeaderColFuzzy(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long, v As String, target As String, found As Long
    FindHeaderColFuzzy = FindHeaderCol(ws, headerName)
    If FindHeaderColFuzzy > 0 Then Exit Function
    target = NormalizeHeader(headerName)
    If Len(target) = 0 Then Exit Function
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        v = NormalizeHeader(CStr(ws.Cells(1, c).Value))
        If v = target Then FindHeaderColFuzzy = c: Exit Function
    Next c
    found = 0
    For c = 1 To lastCol
        v = NormalizeHeader(CStr(ws.Cells(1, c).Value))
        If Len(v) > 0 Then
            If InStr(v, target) > 0 Or InStr(target, v) > 0 Then
                If found = 0 Then found = c Else found = -1: Exit For
            End If
        End If
    Next c
    If found > 0 Then
        FindHeaderColFuzzy = found
        AppendLog "INFO", "RawData", "Колонка '" & headerName & "' сопоставлена с '" & CStr(ws.Cells(1, found).Value) & "' (нестрогий поиск)"
    End If
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, " ", ""): t = Replace(t, "_", ""): t = Replace(t, "-", ""): t = Replace(t, ".", ""): t = Replace(t, ",", "")
    ' латиница -> кириллица (похожие по начертанию буквы)
    t = Replace(t, "P", "Р"): t = Replace(t, "C", "С"): t = Replace(t, "T", "Т"): t = Replace(t, "B", "В")
    t = Replace(t, "H", "Н"): t = Replace(t, "K", "К"): t = Replace(t, "M", "М"): t = Replace(t, "O", "О")
    t = Replace(t, "A", "А"): t = Replace(t, "E", "Е"): t = Replace(t, "X", "Х"): t = Replace(t, "Y", "У")
    NormalizeHeader = t
End Function

' ==========================================================
' Качественная оценка
' ==========================================================

Private Sub EvaluateQualitative(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
                                ByVal pCol As Long, ByVal fCol As Long, ByVal timeCol As Long)
    On Error GoTo EH

    Dim signReq As Integer, t5 As Double, t10 As Double
    Dim r As Long, dP As Double, target5 As Double, target10 As Double
    Dim hit5 As Boolean, hit10 As Boolean, row5 As Long, row10 As Long
    Dim steadyMean As Double, steadyTolMW As Double
    Dim reason As String, failed As String, qStep As String

    qStep = "Проверка включения"
    If Not g.QualEnabled Then
        res.QualPass = True
        res.QualT5Pass = True
        res.QualT10Pass = True
        res.QualSteadyPass = True
        res.T5FactSec = -1
        res.T10FactSec = -1
        res.QualFailedList = ""
        res.QualReason = "Качественная проверка отключена"
        res.PsteadyAvg = 0
        Exit Sub
    End If

    qStep = "Направление требуемого отклика"
    signReq = SgnNZ(res.PReq)
    If signReq = 0 Then
        res.QualPass = True
        res.QualT5Pass = True
        res.QualT10Pass = True
        res.QualSteadyPass = True
        res.T5FactSec = -1
        res.T10FactSec = -1
        res.QualFailedList = ""
        res.QualReason = "Вне зоны отклонения"
        res.PsteadyAvg = 0
        Exit Sub
    End If

    qStep = "Целевые ступени dP5/dP10"
    ' Цели масштабируются к статическому |Pтреб| (Kд = 1, окончание переходного
    ' процесса, п. 4.2.2/4.2.7), ограниченному резервом; dP5/dP10 задают пропорции
    ' (5%/10% = 0.5 = половина Pтреб к моменту t5 - это соответствует Kд(15 с) = 0.5
    ' для ТЭС по Приложению к МУ; полный Pтреб к t10).
    ' Для контрольных испытаний (ступень 10 % Pном) это даёт тот же результат,
    ' т. к. Pтреб = Pном * dp10% / 100.
    Dim pReqAbs As Double, ratio5 As Double, tolMW As Double
    pReqAbs = Abs(res.PReqStatic)
    If pReqAbs < 0.000001 Then pReqAbs = Abs(res.PReq)
    If g.Dp10Pct > 0 Then
        ratio5 = g.Dp5Pct / g.Dp10Pct
    Else
        ratio5 = 0.5
    End If
    tolMW = 0.01 * g.PNom
    target5 = signReq * pReqAbs * ratio5
    target10 = signReq * pReqAbs

    row5 = 0
    row10 = 0
    hit5 = False
    hit10 = False

    qStep = "Поиск t5/t10"
    ' Ступень считается достигнутой с учётом допуска 1 %Pном (п. 4.2.7).
    For r = res.StartRow To res.EndQualRow
        dP = NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0
        If Not hit5 Then
            If signReq * dP >= signReq * target5 - tolMW Then
                hit5 = True
                row5 = r
            End If
        End If
        If Not hit10 Then
            If signReq * dP >= signReq * target10 - tolMW Then
                hit10 = True
                row10 = r
            End If
        End If
    Next r

    ' Если частота к t10 уже вернулась (Pтреб(t) снизился), полный статический Pтреб
    ' объективно не нужен: ступень t10 засчитывается, если в конце качественного окна
    ' нет дефицита относительно текущего Pтреб(t) по Kд(t) (с допуском 1 %Pном).
    If Not hit10 Then
        Dim dFrEnd As Double, preqEnd As Double, haveEnd As Double
        dFrEnd = DeadbandDeviation(NzD(wsRaw.Cells(res.EndQualRow, fCol).Value, st.FNom) - st.FNom, g.Fnch)
        preqEnd = PReqCalc(g.SPct, g.PNom, st.FNom, _
                  DynamicKdByTime(g.EquipType, SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, _
                                                          wsRaw.Cells(res.EndQualRow, timeCol).Value)), dFrEnd, st.MaxReqPct)
        preqEnd = ClampReserve(preqEnd, res.ReservePlus, res.ReserveMinus)
        haveEnd = NzD(wsRaw.Cells(res.EndQualRow, pCol).Value, 0) - res.P0
        If SgnNZ(preqEnd) = 0 Or SgnNZ(preqEnd) <> signReq Then
            ' частота вернулась в зону нечувствительности (или сменила знак) -
            ' требование к t10 снято, если генератор не остался с дефицитом
            hit10 = True
            row10 = res.EndQualRow
        ElseIf signReq * haveEnd >= Abs(preqEnd) - tolMW Then
            hit10 = True
            row10 = res.EndQualRow
        End If
    End If

    qStep = "Расчет фактических времен"
    If hit5 Then
        t5 = SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, wsRaw.Cells(row5, timeCol).Value)
    Else
        t5 = -1
    End If
    If hit10 Then
        t10 = SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, wsRaw.Cells(row10, timeCol).Value)
    Else
        t10 = -1
    End If
    res.T5FactSec = t5
    res.T10FactSec = t10

    qStep = "Среднее установившееся"
    steadyMean = ComputeSteadyMean(wsRaw, pCol, timeCol, res.EndQualRow, st.SteadyWindowSec)
    res.PsteadyAvg = steadyMean
    steadyTolMW = g.SteadyTolPct / 100# * g.PNom

    ' В реальных событиях частота может восстановиться к концу качественного окна.
    ' Поэтому цель установившегося = среднее Pтреб(t) в том же хвостовом окне.
    ' Если частота вернулась, Pтреб_ср близок к нулю и генератор тоже должен вернуться к P0.
    res.PReqSteady = ComputeSteadyPReqMean(wsRaw, fCol, timeCol, res.EndQualRow, _
                                           st.SteadyWindowSec, st.FNom, g.SPct, g.PNom, g.Fnch, _
                                           g.EquipType, wsRaw.Cells(res.StartRow, timeCol).Value, st.MaxReqPct)
    res.PReqSteady = ClampReserve(res.PReqSteady, res.ReservePlus, res.ReserveMinus)

    ' допуск 0.5 с на дискретность выборки
    res.QualT5Pass = (hit5 And t5 <= g.T5Sec + 0.5)
    res.QualT10Pass = (hit10 And t10 <= g.T10Sec + 0.5)
    If g.CheckSteady Then
        res.QualSteadyPass = (Abs((steadyMean - res.P0) - res.PReqSteady) <= steadyTolMW)
    Else
        res.QualSteadyPass = True
    End If

    reason = ""
    failed = ""
    If Not res.QualT5Pass Then
        reason = reason & "Не достигнута 1-я ступень (" & Format(ratio5 * 100#, "0") & " % Pтреб) к t5=" & g.T5Sec & "c; "
        failed = failed & "t5; "
    End If
    If Not res.QualT10Pass Then
        reason = reason & "Не достигнут Pтреб к t10=" & g.T10Sec & "c; "
        failed = failed & "t10; "
    End If
    If g.CheckSteady And (Not res.QualSteadyPass) Then
        reason = reason & "Установившееся отклонение от Pтреб_ср (" & Format(res.PReqSteady, "0.000") _
                        & ") выходит за допуск ±" & Format(steadyTolMW, "0.000") & " МВт; "
        failed = failed & "уст; "
    End If
    If Not g.CheckSteady Then reason = reason & "Контроль установившегося отключен; "

    res.QualPass = (res.QualT5Pass And res.QualT10Pass And res.QualSteadyPass)
    If Len(failed) > 0 Then
        If Right$(failed, 2) = "; " Then failed = Left$(failed, Len(failed) - 2)
        res.QualFailedList = failed
    Else
        res.QualFailedList = ""
    End If

    If res.QualPass Then
        res.QualReason = "Качественно: ОК; t5=" & IIf(t5 >= 0, Format(t5, "0.0"), "н/д") & "с; t10=" & IIf(t10 >= 0, Format(t10, "0.0"), "н/д") & "с"
    Else
        res.QualReason = reason & "t5=" & IIf(t5 >= 0, Format(t5, "0.0"), "н/д") & "с; t10=" & IIf(t10 >= 0, Format(t10, "0.0"), "н/д") & "с"
    End If
    Exit Sub

EH:
    Err.Raise vbObjectError + 2299, , "EvaluateQualitative (" & g.Generator & " / " & qStep & "): " & Err.Description
End Sub

Private Function ComputeSteadyMean(ByVal wsRaw As Worksheet, ByVal pCol As Long, ByVal timeCol As Long, _
                                    ByVal endRow As Long, ByVal windowSec As Double) As Double
    Dim startRow As Long, r As Long
    Dim sumP As Double, cnt As Long
    If windowSec <= 0 Then windowSec = 30
    startRow = RowByTimeOffset(wsRaw, timeCol, endRow, -windowSec)
    If startRow < 2 Then startRow = 2
    If startRow >= endRow Then startRow = endRow
    For r = startRow To endRow
        If IsNumeric(wsRaw.Cells(r, pCol).Value) Then
            sumP = sumP + CDbl(wsRaw.Cells(r, pCol).Value)
            cnt = cnt + 1
        End If
    Next r
    If cnt > 0 Then ComputeSteadyMean = sumP / cnt Else ComputeSteadyMean = 0
End Function

Private Function ComputeSteadyPReqMean(ByVal wsRaw As Worksheet, ByVal freqCol As Long, ByVal timeCol As Long, _
                                        ByVal endRow As Long, ByVal windowSec As Double, _
                                        ByVal fNom As Double, ByVal sPct As Double, ByVal pNom As Double, _
                                        ByVal fnch As Double, _
                                        ByVal equipType As String, ByVal tStart As Variant, _
                                        ByVal maxReqPct As Double) As Double
    ' Усреднённый Pтреб(t) за окно [endRow-windowSec ; endRow].
    ' Pтреб(t) = -100/S * Pном/fном * Kd * dFr(t), где dFr - отклонение за fнч (с ограничением maxReqPct).
    Dim startRow As Long, r As Long
    Dim sumP As Double, cnt As Long
    Dim dFr As Double, fv As Double, kdEff As Double, tSec As Double
    If windowSec <= 0 Then windowSec = 30
    If sPct <= 0 Then Exit Function
    startRow = RowByTimeOffset(wsRaw, timeCol, endRow, -windowSec)
    If startRow < 2 Then startRow = 2
    If startRow >= endRow Then startRow = endRow
    For r = startRow To endRow
        If IsNumeric(wsRaw.Cells(r, freqCol).Value) Then
            fv = CDbl(wsRaw.Cells(r, freqCol).Value) - fNom
            dFr = DeadbandDeviation(fv, fnch)
            tSec = SecBetween(tStart, wsRaw.Cells(r, timeCol).Value)
            kdEff = DynamicKdByTime(equipType, tSec)
            sumP = sumP + PReqCalc(sPct, pNom, fNom, kdEff, dFr, maxReqPct)
            cnt = cnt + 1
        End If
    Next r
    If cnt > 0 Then ComputeSteadyPReqMean = sumP / cnt Else ComputeSteadyPReqMean = 0
End Function

' ==========================================================
' Характер переходного процесса: Монотонный / Апериодический / Колебательный
' ==========================================================

Private Sub EvaluateTransient(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult, _
                              ByVal pCol As Long, ByVal timeCol As Long)
    Dim r As Long, n As Long, i As Long
    Dim arr() As Double
    Dim extrCount As Long
    Dim prevDir As Integer, curDir As Integer
    Dim dP As Double, prev As Double
    Dim amps() As Double, ampCount As Long
    Dim lastExtrVal As Double
    Dim hasLast As Boolean
    Dim maxAbs As Double, noise As Double
    Dim smoothed() As Double

    If res.EndQualRow <= res.StartRow Then
        res.TransientType = "н/д"
        res.NumExtrema = 0
        Exit Sub
    End If

    n = res.EndQualRow - res.StartRow + 1
    ReDim arr(1 To n)
    i = 0
    For r = res.StartRow To res.EndQualRow
        i = i + 1
        arr(i) = NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0
        If Abs(arr(i)) > maxAbs Then maxAbs = Abs(arr(i))
    Next r

    If n < 5 Or maxAbs < 0.0001 Then
        res.TransientType = "н/д"
        res.NumExtrema = 0
        Exit Sub
    End If

    ' Сглаживание окном 3 точки (убираем шум дискретизации)
    ReDim smoothed(1 To n)
    smoothed(1) = arr(1)
    smoothed(n) = arr(n)
    For i = 2 To n - 1
        smoothed(i) = (arr(i - 1) + arr(i) + arr(i + 1)) / 3#
    Next i

    noise = maxAbs * 0.05   ' Порог "значимого" экстремума: 5 % от амплитуды
    extrCount = 0
    ReDim amps(1 To n)
    ampCount = 0
    prevDir = 0
    lastExtrVal = smoothed(1)
    hasLast = False

    For i = 2 To n - 1
        dP = smoothed(i) - smoothed(i - 1)
        If dP > noise Then
            curDir = 1
        ElseIf dP < -noise Then
            curDir = -1
        Else
            curDir = 0
        End If

        If prevDir <> 0 And curDir <> 0 And prevDir <> curDir Then
            ' Локальный экстремум в точке i-1
            If Abs(smoothed(i - 1) - lastExtrVal) >= noise Or Not hasLast Then
                extrCount = extrCount + 1
                ampCount = ampCount + 1
                amps(ampCount) = Abs(smoothed(i - 1) - lastExtrVal)
                lastExtrVal = smoothed(i - 1)
                hasLast = True
            End If
        End If
        If curDir <> 0 Then prevDir = curDir
    Next i

    res.NumExtrema = extrCount

    Dim decaying As Boolean
    decaying = True
    For i = 2 To ampCount
        If amps(i) > amps(i - 1) * 1.1 Then
            decaying = False
            Exit For
        End If
    Next i

    If extrCount <= 1 Then
        res.TransientType = "Монотонный"
    ElseIf extrCount <= 3 And decaying Then
        res.TransientType = "Апериодический"
    Else
        res.TransientType = "Колебательный"
    End If
End Sub

' ==========================================================
' Запись листа данных генератора
' ==========================================================

Private Sub WriteGeneratorSheet(ByVal wsRaw As Worksheet, ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim ws As Worksheet
    Dim shName As String
    Dim timeCol As Long, pCol As Long, fCol As Long
    Dim endRow As Long, r As Long, outR As Long
    Dim dP As Double, dFr As Double
    Dim displayStartRow As Long, chartEndRow As Long
    Dim target5Val As Double, target10Val As Double
    Dim targetPreq As Double
    Dim tolPreq As Double
    Dim signReq As Integer

    timeCol = FindHeaderCol(wsRaw, "Время")
    pCol = FindHeaderColFuzzy(wsRaw, g.PowerHeader)
    fCol = FindHeaderColFuzzy(wsRaw, g.FreqHeader)

    shName = GeneratorSheetName(g, res.EventIdx)
    Set ws = EnsureSheet(shName)
    ws.Cells.Clear
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop

    ' Шапка (колонки A:B и D:E)
    Dim quantStatusStr As String
    If res.Limited Then
        quantStatusStr = "Ограничен " & res.LimitType
    ElseIf res.QuantPass Then
        quantStatusStr = "ОК"
    Else
        quantStatusStr = "Нарушение"
    End If

    ws.Range("A1:B1").Value = Array("Станция", g.Station)
    ws.Range("A2:B2").Value = Array("Генератор", g.Generator)
    ws.Range("A3:B3").Value = Array("Тип", g.EquipType)
    ws.Range("A4:B4").Value = Array("Старт (расч.)", res.StartTime)
    ws.Range("A5:B5").Value = Array("Выход за fнч", res.FirstExceedTime)
    ws.Range("A6:B6").Value = Array("Колич. статус", IIf(res.QuantApplicable, quantStatusStr, "Н/П (п. 4.1.2)"))
    ws.Range("A7:B7").Value = Array("Кач. статус", IIf(res.QualApplicable, IIf(res.QualPass, "ОК", "Нарушение"), "Н/П (п. 4.1.3)"))
    ws.Range("A8:B8").Value = Array("Характер", res.TransientType)
    ws.Range("A9:B9").Value = Array("Амплитуда, %Pном", Round(res.AmplPctPnom, 2))
    ws.Cells(1, 4).Resize(1, 2).Value = Array("P0, МВт", res.P0)
    ws.Cells(2, 4).Resize(1, 2).Value = Array("Pтек(tр), МВт", res.PTek)
    ws.Cells(3, 4).Resize(1, 2).Value = Array("Pуст_сред, МВт", res.PsteadyAvg)
    ws.Cells(4, 4).Resize(1, 2).Value = Array("dF макс, Гц", res.Df)
    ws.Cells(5, 4).Resize(1, 2).Value = Array("dFr(tр), Гц", res.Dfr)
    ws.Cells(6, 4).Resize(1, 2).Value = Array("Pтреб(tр), МВт", res.PReq)
    ws.Cells(7, 4).Resize(1, 2).Value = Array("Pфакт(tр), МВт", res.PFact)
    ws.Cells(8, 4).Resize(1, 2).Value = Array("Колич. %", res.QuantPct)
    ws.Cells(9, 4).Resize(1, 2).Value = Array("Экстремумов", res.NumExtrema)

    ' Блок 1.7.0 (X:Y): событие, tр, применимость, тип реакции, колебания, ТИ, h(15 с)
    ws.Cells(1, 24).Resize(1, 2).Value = Array("Событие №", res.EventIdx)
    ws.Cells(2, 24).Resize(1, 2).Value = Array("tр (п. 4.2.5)", res.TrTime)
    ws.Cells(3, 24).Resize(1, 2).Value = Array("tр, с от старта", res.TrSec)
    ws.Cells(4, 24).Resize(1, 2).Value = Array("Режим в tр", IIf(res.TrQuasiSteady, "квазиустановившийся (±15 с, 10 мГц)", "переходный"))
    ws.Cells(5, 24).Resize(1, 2).Value = Array("Применимость", res.Applicability)
    ws.Cells(6, 24).Resize(1, 2).Value = Array("Тип реакции (п. 4.4.1)", IIf(Len(res.ReactionType) > 0, res.ReactionType, "-"))
    ws.Cells(7, 24).Resize(1, 2).Value = Array("Колебания (п. 4.2.14.4)", IIf(Len(res.OscNote) > 0, res.OscNote, "нет"))
    ws.Cells(8, 24).Resize(1, 2).Value = Array("Качество ТИ", IIf(Len(res.TelemetryNote) > 0, res.TelemetryNote, "ОК"))
    ws.Cells(9, 24).Resize(1, 2).Value = Array("h(15 с), Прил. п. 6.2", IIf(res.H15 > -900, Round(res.H15, 3), res.H15Note))
    ws.Cells(10, 24).Resize(1, 2).Value = Array("Pтреб/Pфакт (Tинт), МВт", Format(res.PReqRef, "0.000") & " / " & Format(res.PFactRef, "0.000"))
    ws.Cells(11, 24).Resize(1, 2).Value = Array("Режим / РМ", st.Mode & " / " & IIf(g.HasRM, "с РМ", "без РМ") & IIf(g.ForcedNoRM, " (S=6 %, fнч=0.15 Гц)", "") & " / " & g.Layout)

    ws.Cells(1, 7).Resize(1, 2).Value = Array("Pmax, МВт", res.PMaxEff)
    ws.Cells(2, 7).Resize(1, 2).Value = Array("Pmin, МВт", res.PMinEff)
    ws.Cells(3, 7).Resize(1, 2).Value = Array("Резерв '+', МВт", res.ReservePlus)
    ws.Cells(4, 7).Resize(1, 2).Value = Array("Резерв '-', МВт", res.ReserveMinus)
    ws.Cells(5, 7).Resize(1, 2).Value = Array("Pтреб исх., МВт", res.PReqOrig)
    ws.Cells(6, 7).Resize(1, 2).Value = Array("Ограничение", IIf(res.Limited, "Да (" & res.LimitType & ")", "нет"))
    ws.Cells(7, 7).Resize(1, 2).Value = Array("Kд (колич.), факт", res.KdUsedQuant)
    ws.Cells(8, 7).Resize(1, 2).Value = Array("Профиль Kд(t)", res.KdProfile)
    ws.Cells(9, 7).Resize(1, 2).Value = Array("Реакция 15 с (п. 4.2.11)", res.React15Status & " (dP15 = " & Format(res.DP15, "0.000") & " МВт)")
    ws.Cells(10, 7).Resize(1, 2).Value = Array("Дефицит (п. 4.2.14.1)", _
        IIf(res.DeficitSec > 0, "до " & Format(res.DeficitMaxPct, "0.00") & " %Pном, " & Format(res.DeficitSec, "0") & " с", "нет"))

    ws.Range("A11:W11").Value = Array( _
        "Время", "Частота, Гц", "P, МВт", "dPфакт, МВт", "Pтреб_накоп, МВт", "dFr, Гц", _
        "Маркер tр", "Уровень +допуск", "Уровень -допуск", _
        "Маркер t5", "Маркер t10", "Маркер выхода за fнч", _
        "dPmax", "dPmin", _
        "Pтреб_абс, МВт", "Уровень +допуск_абс", "Уровень -допуск_абс", _
        "Pmax, МВт", "Pmin, МВт", _
        "Маркер t5_абс", "Маркер t10_абс", "Маркер fнч_абс", "Маркер tр_абс" _
    )

    If st.PreBufferSec > 0 Then
        displayStartRow = RowByTimeOffset(wsRaw, timeCol, res.StartRow, -st.PreBufferSec)
        If displayStartRow < 2 Then displayStartRow = 2
        If displayStartRow > res.StartRow Then displayStartRow = res.StartRow
    Else
        displayStartRow = res.StartRow
    End If
    endRow = RowByTimeOffset(wsRaw, timeCol, res.StartRow, MaxD(st.QuantIntervalSec, g.T10Sec))
    If res.EventEndRow > endRow Then endRow = res.EventEndRow
    If res.TrRow > endRow Then endRow = res.TrRow
    outR = 12

    ' Уровни для маркеров. Целевой уровень = Pтреб (масштабируется к событию).
    signReq = SgnNZ(res.PReq)
    targetPreq = res.PReq
    tolPreq = g.PNom * 0.01
    ' Для вертикальных маркеров подбираем достаточный размах, чтобы линия была видна.
    ' Берём не меньше размаха Pmax/Pmin, чтобы маркер пересекал обе границы.
    Dim markerSpan As Double, markerSpanAbs As Double
    Dim dPmaxRel As Double, dPminRel As Double
    Dim markerHiAbs As Double, markerLoAbs As Double
    dPmaxRel = res.PMaxEff - res.P0
    dPminRel = res.PMinEff - res.P0
    markerSpan = MaxD(Abs(res.PReq) * 1.5, g.PNom * g.Dp10Pct / 100#)
    markerSpan = MaxD(markerSpan, MaxD(Abs(dPmaxRel), Abs(dPminRel)))
    If markerSpan <= 0 Then markerSpan = g.PNom * 0.1
    markerSpanAbs = markerSpan
    markerHiAbs = res.P0 + markerSpanAbs
    markerLoAbs = res.P0 - markerSpanAbs
    target10Val = targetPreq

    Dim preqStep As Double
    For r = displayStartRow To endRow
        dP = NzD(wsRaw.Cells(r, pCol).Value, 0) - res.P0
        dFr = DeadbandDeviation(NzD(wsRaw.Cells(r, fCol).Value, st.FNom) - st.FNom, g.Fnch)
        preqStep = ClampReserve(PReqCalc(g.SPct, g.PNom, st.FNom, DynamicKdByTime(g.EquipType, _
                   SecBetween(wsRaw.Cells(res.StartRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)), dFr, st.MaxReqPct), _
                   res.ReservePlus, res.ReserveMinus)

        ws.Cells(outR, 1).Value = wsRaw.Cells(r, timeCol).Value
        ws.Cells(outR, 2).Value = wsRaw.Cells(r, fCol).Value
        ws.Cells(outR, 3).Value = wsRaw.Cells(r, pCol).Value
        ws.Cells(outR, 4).Value = dP
        ws.Cells(outR, 5).Value = preqStep
        ws.Cells(outR, 6).Value = dFr
        ws.Cells(outR, 8).Value = targetPreq + tolPreq
        ws.Cells(outR, 9).Value = targetPreq - tolPreq
        ws.Cells(outR, 13).Value = dPmaxRel
        ws.Cells(outR, 14).Value = dPminRel
        ws.Cells(outR, 15).Value = res.P0 + preqStep
        ws.Cells(outR, 16).Value = res.P0 + targetPreq + tolPreq
        ws.Cells(outR, 17).Value = res.P0 + targetPreq - tolPreq
        ws.Cells(outR, 18).Value = res.PMaxEff
        ws.Cells(outR, 19).Value = res.PMinEff
        outR = outR + 1
    Next r

    ' Маркеры t5 / t10 / выхода за fнч - две точки на линии (нижний и верхний уровень)
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T5Sec, 10, markerSpan, -markerSpan
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T10Sec, 11, markerSpan, -markerSpan
    FillMarkerColumnAtTime ws, displayStartRow, endRow, wsRaw, timeCol, res.FirstExceedTime, 12, markerSpan, -markerSpan
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T5Sec, 20, markerHiAbs, markerLoAbs
    FillMarkerColumn ws, displayStartRow, endRow, wsRaw, timeCol, res.StartRow, g.T10Sec, 21, markerHiAbs, markerLoAbs
    FillMarkerColumnAtTime ws, displayStartRow, endRow, wsRaw, timeCol, res.FirstExceedTime, 22, markerHiAbs, markerLoAbs
    FillMarkerColumnAtTime ws, displayStartRow, endRow, wsRaw, timeCol, res.TrTime, 7, markerSpan, -markerSpan
    FillMarkerColumnAtTime ws, displayStartRow, endRow, wsRaw, timeCol, res.TrTime, 23, markerHiAbs, markerLoAbs

    chartEndRow = outR - 1

    ws.Columns("A:W").AutoFit
    ApplyGeneratorSheetFormats ws
    ws.Columns("X").ColumnWidth = 24
    ws.Columns("Y").ColumnWidth = 60
    ' Вывод записываем после AutoFit, чтобы длинный текст не растягивал колонку B
    ws.Cells(10, 1).Value = "Вывод"
    ws.Cells(10, 2).Value = res.Verdict
    ws.Cells(10, 1).Font.Bold = True
End Sub

Private Sub FillMarkerColumn(ByVal ws As Worksheet, ByVal rr1 As Long, ByVal rr2 As Long, _
                             ByVal wsRaw As Worksheet, ByVal timeCol As Long, _
                             ByVal startRow As Long, ByVal offsetSec As Double, _
                             ByVal outCol As Long, ByVal vHi As Double, ByVal vLo As Double)
    Dim mkRow As Long
    mkRow = RowByTimeOffset(wsRaw, timeCol, startRow, offsetSec)
    If mkRow < 2 Then Exit Sub
    FillMarkerColumnAtTime ws, rr1, rr2, wsRaw, timeCol, wsRaw.Cells(mkRow, timeCol).Value, outCol, vHi, vLo
End Sub

Private Sub FillMarkerColumnAtTime(ByVal ws As Worksheet, ByVal rr1 As Long, ByVal rr2 As Long, _
                                   ByVal wsRaw As Worksheet, ByVal timeCol As Long, _
                                   ByVal tMark As Variant, ByVal outCol As Long, _
                                   ByVal vHi As Double, ByVal vLo As Double)
    Dim r As Long, rowIdx As Long, tM As Double
    Dim tPrev As Double, tCur As Double, match As Boolean
    If Not IsDate(tMark) Then Exit Sub
    tM = CDbl(CDate(tMark))
    rowIdx = 12
    tPrev = 0
    match = False
    For r = rr1 To rr2
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then
            tCur = CDbl(CDate(wsRaw.Cells(r, timeCol).Value))
            If (tPrev <= tM And tCur >= tM) Or Abs(tCur - tM) < 0.5 / 86400# Then
                ws.Cells(rowIdx, outCol).Value = vHi
                ws.Cells(rowIdx + 1, outCol).Value = vLo
                match = True
                Exit For
            End If
            tPrev = tCur
        End If
        rowIdx = rowIdx + 1
    Next r
    If Not match Then Exit Sub
End Sub

Private Sub ApplyGeneratorSheetFormats(ByVal ws As Worksheet)
    ws.Range("A1:A9").NumberFormat = "@"
    ws.Range("B1:B3").NumberFormat = "@"
    ws.Range("B4:B5").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B6:B8").NumberFormat = "@"
    ws.Range("B9").NumberFormat = "0.00"
    ws.Range("D1:D9").NumberFormat = "@"
    ws.Range("E1:E9").NumberFormat = "0.000"
    ws.Range("A10:B10").NumberFormat = "@"
    ws.Range("G1:G10").NumberFormat = "@"
    ws.Range("H1:H5").NumberFormat = "0.000"
    ws.Range("H6").NumberFormat = "@"
    ws.Range("H7").NumberFormat = "0.000"
    ws.Range("H8:H10").NumberFormat = "@"
    ws.Range("A11:W11").NumberFormat = "@"
    ws.Range("A12:A100000").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B12:W100000").NumberFormat = "0.000"
    ws.Range("X1:X11").NumberFormat = "@"
    ws.Range("Y1").NumberFormat = "0"
    ws.Range("Y2").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("Y3").NumberFormat = "0"
    ws.Range("Y4:Y11").NumberFormat = "@"
    ws.Range("Y4:Y11").WrapText = True
End Sub

' ==========================================================
' Лист с графиком
' ==========================================================

Private Sub WriteGeneratorChartSheet(ByRef st As TSettings, ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim wsData As Worksheet, wsChart As Worksheet
    Dim dataSheet As String, chartSheet As String
    Dim lastDataRow As Long
    Dim chartObj As ChartObject
    Dim startDataRow As Long, endChartDataRow As Long
    Dim ySpan As Double
    Dim yMin As Double, yMax As Double, yPad As Double
    Dim farPos As Double, farNeg As Double, farAbs As Double

    dataSheet = GeneratorSheetName(g, res.EventIdx)
    chartSheet = GeneratorChartSheetName(g, res.EventIdx)
    Set wsData = ThisWorkbook.Worksheets(dataSheet)
    Set wsChart = EnsureSheet(chartSheet)

    wsChart.Cells.Clear
    Do While wsChart.ChartObjects.Count > 0
        wsChart.ChartObjects(1).Delete
    Loop

    ' Заголовок
    wsChart.Range("A1").Value = "График ОПРЧ: " & g.Station & " / " & g.Generator & IIf(res.EventIdx > 1, " (событие " & res.EventIdx & ")", "")
    wsChart.Range("A1").Font.Bold = True
    wsChart.Range("A1").Font.Size = 14

    ' Блок "Вывод"
    WriteChartVerdictBlock wsChart, g, res, st

    lastDataRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastDataRow < 12 Then Exit Sub

    startDataRow = 12
    endChartDataRow = FindChartEndRow(wsData, 12, lastDataRow, res.StartTime, st.ChartIntervalSec)
    If endChartDataRow < startDataRow Then endChartDataRow = lastDataRow

    Set chartObj = wsChart.ChartObjects.Add(10, 215, 1020, 360)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Мониторинг ОПРЧ: " & g.Station & " / " & g.Generator & IIf(res.EventIdx > 1, " (событие " & res.EventIdx & ")", "")
        On Error Resume Next
        .Axes(xlCategory).CategoryType = xlCategoryScale
        On Error GoTo 0
    End With

    ' Основные ряды
    If ChartOpt("REL_PFACT", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 4, "Pфакт, МВт", False
    If ChartOpt("REL_PREQ", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 5, "Pтреб, МВт", False
    If ChartOpt("REL_TOL_P", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 8, "+Допуск уст.", False
    If ChartOpt("REL_TOL_M", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 9, "-Допуск уст.", False
    If ChartOpt("REL_LIMITS", True) Then
        AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 13, "Pmax (dPmax)"
        AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 14, "Pmin (dPmin)"
    End If
    If ChartOpt("REL_FREQ", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 2, "Частота, Гц", True

    ' Вертикальные маркеры t5 / t10 / выхода за fнч - через отдельные "точечные" ряды
    If ChartOpt("REL_MARKERS", True) Then
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 10, "t5"
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 11, "t10"
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 12, "Выход за fнч"
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 7, "tр"
    End If

    On Error Resume Next
    GetRangeMinMaxByCols wsData, startDataRow, endChartDataRow, Array(4, 5, 8, 9, 13, 14), yMin, yMax
    farPos = MaxAbsInCols(wsData, startDataRow, endChartDataRow, Array(4, 5), True)
    farNeg = MaxAbsInCols(wsData, startDataRow, endChartDataRow, Array(4, 5), False)
    farAbs = MaxD(farPos, farNeg)
    If farAbs > 0 Then
        If farPos >= farNeg Then
            yMax = MaxD(yMax, farPos * 1.05)
        Else
            yMin = -MaxD(Abs(yMin), farNeg * 1.05)
        End If
    End If
    yPad = MaxD((yMax - yMin) * 0.1, g.PNom * 0.01)
    If yPad <= 0 Then yPad = 0.5
    chartObj.Chart.Axes(xlValue, xlPrimary).MinimumScale = yMin - yPad
    chartObj.Chart.Axes(xlValue, xlPrimary).MaximumScale = yMax + yPad
    chartObj.Chart.Axes(xlCategory).TickLabels.NumberFormat = "hh:mm:ss"
    With chartObj.Chart.Axes(xlValue, xlSecondary)
        ySpan = MaxD(Abs(res.Df) * 1.2, 2# * g.Fnch)
        If ySpan < 0.1 Then ySpan = 0.1
        .MinimumScale = st.FNom - ySpan
        .MaximumScale = st.FNom + ySpan
    End With
    On Error GoTo 0
    ApplyChartAxesFormat chartObj.Chart, wsData, startDataRow, endChartDataRow, "dP, МВт", "f, Гц"

    ' Второй график: абсолютные единицы (P, МВт и f, Гц)
    Set chartObj = wsChart.ChartObjects.Add(10, 595, 1020, 360)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Мониторинг ОПРЧ (абсолютные): " & g.Station & " / " & g.Generator & IIf(res.EventIdx > 1, " (событие " & res.EventIdx & ")", "")
        On Error Resume Next
        .Axes(xlCategory).CategoryType = xlCategoryScale
        On Error GoTo 0
    End With

    If ChartOpt("ABS_PFACT", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 3, "Pфакт, МВт", False
    If ChartOpt("ABS_PREQ", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 15, "Pтреб, МВт (абс)", False
    If ChartOpt("ABS_TOL_P", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 16, "+Допуск уст., МВт (абс)", False
    If ChartOpt("ABS_TOL_M", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 17, "-Допуск уст., МВт (абс)", False
    If ChartOpt("ABS_LIMITS", True) Then
        AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 18, "Pmax, МВт"
        AddLimitLineSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 19, "Pmin, МВт"
    End If
    If ChartOpt("ABS_FREQ", True) Then AddSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 2, "Частота, Гц", True
    If ChartOpt("ABS_MARKERS", False) Then
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 20, "t5 abs"
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 21, "t10 abs"
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 22, "fнч abs"
        AddMarkerSeries chartObj.Chart, wsData, startDataRow, endChartDataRow, 1, 23, "tр abs"
    End If

    On Error Resume Next
    GetRangeMinMaxByCols wsData, startDataRow, endChartDataRow, Array(3, 15, 16, 17, 18, 19), yMin, yMax
    farPos = MaxAbsInCols(wsData, startDataRow, endChartDataRow, Array(3, 15), True)
    farNeg = MaxAbsInCols(wsData, startDataRow, endChartDataRow, Array(3, 15), False)
    farAbs = MaxD(farPos, farNeg)
    If farAbs > 0 Then
        If farPos >= farNeg Then
            yMax = MaxD(yMax, farPos * 1.05)
        Else
            yMin = -MaxD(Abs(yMin), farNeg * 1.05)
        End If
    End If
    yPad = MaxD((yMax - yMin) * 0.08, g.PNom * 0.005)
    If yPad <= 0 Then yPad = 1
    chartObj.Chart.Axes(xlValue, xlPrimary).MinimumScale = yMin - yPad
    chartObj.Chart.Axes(xlValue, xlPrimary).MaximumScale = yMax + yPad
    chartObj.Chart.Axes(xlCategory).TickLabels.NumberFormat = "hh:mm:ss"
    With chartObj.Chart.Axes(xlValue, xlSecondary)
        ySpan = MaxD(Abs(res.Df) * 1.2, 2# * g.Fnch)
        If ySpan < 0.1 Then ySpan = 0.1
        .MinimumScale = st.FNom - ySpan
        .MaximumScale = st.FNom + ySpan
    End With
    On Error GoTo 0
    ApplyChartAxesFormat chartObj.Chart, wsData, startDataRow, endChartDataRow, "P, МВт", "f, Гц"
End Sub

' Оформление осей по И.2 Приложения 1: подписи осей с единицами измерения,
' вертикальная сетка не реже чем через 15 с (по числу точек на 15 с), подписи времени.
Private Sub ApplyChartAxesFormat(ByVal ch As Chart, ByVal wsData As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                                 ByVal yTitle As String, ByVal y2Title As String)
    Dim n As Long, spanSec As Double, ptsPer15 As Long
    On Error Resume Next
    n = r2 - r1 + 1
    spanSec = SecBetween(wsData.Cells(r1, 1).Value, wsData.Cells(r2, 1).Value)
    If spanSec > 0 And n > 1 Then
        ptsPer15 = CLng(15# * (n - 1) / spanSec)
    Else
        ptsPer15 = 15
    End If
    If ptsPer15 < 1 Then ptsPer15 = 1
    With ch.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "Время, чч:мм:сс"
        .HasMajorGridlines = True
        .MajorGridlines.Format.Line.ForeColor.RGB = RGB(217, 217, 217)
        .TickMarkSpacing = ptsPer15
        .TickLabelSpacing = ptsPer15
        .TickLabels.Orientation = 90
    End With
    With ch.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = yTitle
        .HasMajorGridlines = True
    End With
    With ch.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = y2Title
    End With
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom
    On Error GoTo 0
End Sub

Private Sub GetRangeMinMaxByCols(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                                 ByVal cols As Variant, ByRef outMin As Double, ByRef outMax As Double)
    Dim i As Long, r As Long, c As Long, v As Variant
    Dim inited As Boolean
    inited = False
    For i = LBound(cols) To UBound(cols)
        c = CLng(cols(i))
        For r = r1 To r2
            v = ws.Cells(r, c).Value
            If IsNumeric(v) Then
                If Not inited Then
                    outMin = CDbl(v)
                    outMax = CDbl(v)
                    inited = True
                Else
                    If CDbl(v) < outMin Then outMin = CDbl(v)
                    If CDbl(v) > outMax Then outMax = CDbl(v)
                End If
            End If
        Next r
    Next i
    If Not inited Then
        outMin = 0
        outMax = 1
    End If
    If outMax < outMin Then
        outMax = outMin + 1
    ElseIf Abs(outMax - outMin) < 0.000001 Then
        outMax = outMax + 1
        outMin = outMin - 1
    End If
End Sub

Private Function MaxAbsInCols(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                              ByVal cols As Variant, ByVal positiveSide As Boolean) As Double
    Dim i As Long, r As Long, c As Long
    Dim v As Variant, vv As Double
    For i = LBound(cols) To UBound(cols)
        c = CLng(cols(i))
        For r = r1 To r2
            v = ws.Cells(r, c).Value
            If IsNumeric(v) Then
                vv = CDbl(v)
                If positiveSide Then
                    If vv > MaxAbsInCols Then MaxAbsInCols = vv
                Else
                    If -vv > MaxAbsInCols Then MaxAbsInCols = -vv
                End If
            End If
        Next r
    Next i
End Function

Private Sub WriteChartVerdictBlock(ByVal wsChart As Worksheet, ByRef g As TGenCfg, ByRef res As TGenResult, ByRef st As TSettings)
    Dim quantCell As String
    If res.Limited Then
        quantCell = "Ограничен " & res.LimitType & " (" & Format(res.QuantPct, "0") & " % от доступного)"
    ElseIf res.QuantPass Then
        quantCell = "ОК (" & Format(res.QuantPct, "0") & " %)"
    Else
        quantCell = "Нарушение (" & Format(res.QuantPct, "0") & " %)"
    End If

    wsChart.Range("A3").Value = "Количественный:"
    wsChart.Range("B3").Value = quantCell
    wsChart.Range("D3").Value = "Амплитуда:"
    wsChart.Range("E3").Value = Format(res.AmplPctPnom, "0.0") & " %Pном" & IIf(Len(res.AmplitudeTag) > 0, " / " & res.AmplitudeTag, "")

    wsChart.Range("A4").Value = "Качественный:"
    wsChart.Range("B4").Value = IIf(res.QualPass, "ОК", "Нарушение")
    wsChart.Range("D4").Value = "t5 / t10:"
    wsChart.Range("E4").Value = FormatSecOrNA(res.T5FactSec) & " / " & FormatSecOrNA(res.T10FactSec) & " с"

    wsChart.Range("A5").Value = "Характер:"
    wsChart.Range("B5").Value = res.TransientType & " (экстремумов " & res.NumExtrema & ")"
    wsChart.Range("D5").Value = "Перерегулирование:"
    wsChart.Range("E5").Value = IIf(res.Overshoot, "Да", "нет")

    wsChart.Range("A6").Value = "Уст_сред, МВт:"
    wsChart.Range("B6").Value = Format(res.PsteadyAvg, "0.000") & " (цель " & Format(res.P0 + res.PReqSteady, "0.000") & ")"
    wsChart.Range("D6").Value = "Проваленные подп.:"
    wsChart.Range("E6").Value = IIf(Len(res.QualFailedList) > 0, res.QualFailedList, "-")

    wsChart.Range("A7").Value = "Pmax / Pmin, МВт:"
    wsChart.Range("B7").Value = Format(res.PMaxEff, "0.0") & " / " & Format(res.PMinEff, "0.0") & _
                                " (резерв +" & Format(res.ReservePlus, "0.0") & " / -" & _
                                Format(res.ReserveMinus, "0.0") & ")"
    wsChart.Range("D7").Value = "Pтреб (исх / прим.):"
    wsChart.Range("E7").Value = Format(res.PReqOrig, "0.000") & " / " & Format(res.PReq, "0.000") & _
                                IIf(res.Limited, " (ограничен " & res.LimitType & ")", "")

    wsChart.Range("A8").Value = "Kд(t):"
    wsChart.Range("B8").Value = Format(res.KdUsedQuant, "0.000") & " / " & res.KdProfile

    wsChart.Range("A9").Value = "Реакция 15 с:"
    wsChart.Range("B9").Value = res.React15Status & " (dP15 = " & Format(res.DP15, "0.000") & " МВт; п. 4.2.11)"
    wsChart.Range("D9").Value = "Дефицит:"
    wsChart.Range("E9").Value = IIf(res.DeficitSec > 0, "до " & Format(res.DeficitMaxPct, "0.00") & " %Pном, " & _
                                Format(res.DeficitSec, "0") & " с (п. 4.2.14.1)", "нет")

    wsChart.Range("A10").Value = "tр / применимость:"
    wsChart.Range("B10").Value = IIf(IsDate(res.TrTime), Format(res.TrTime, "hh:mm:ss"), "-") & " (" & Format(res.TrSec, "0") & " с, " & _
                                 IIf(res.TrQuasiSteady, "квазиуст.", "переходный") & "); " & res.Applicability
    wsChart.Range("A11").Value = "Тип реакции / ТИ:"
    wsChart.Range("B11").Value = IIf(Len(res.ReactionType) > 0, res.ReactionType & " (п. 4.4.1)", "-") & _
                                 IIf(Len(res.QualClauseVerdict) > 0, "; " & res.QualClauseVerdict, "") & _
                                 IIf(Len(res.TelemetryNote) > 0, "; " & res.TelemetryNote, "") & _
                                 IIf(Len(res.OscNote) > 0, "; " & res.OscNote, "") & _
                                 IIf(res.H15 > -900, "; h(15 с) = " & Format(res.H15, "0.00"), "")

    wsChart.Range("A12").Value = "Вывод:"
    wsChart.Range("B12").Value = res.Verdict

    wsChart.Range("A3:A12").Font.Bold = True
    wsChart.Range("D3:D9").Font.Bold = True
    wsChart.Range("A3:E12").NumberFormat = "@"
    If res.Limited Then
        wsChart.Range("B3").Font.Color = RGB(156, 87, 0)
        wsChart.Range("B3").Interior.Color = RGB(255, 235, 156)
        wsChart.Range("E7").Font.Color = RGB(156, 87, 0)
    ElseIf Not res.QuantPass And res.QuantApplicable Then
        wsChart.Range("B3").Font.Color = RGB(192, 0, 0)
    End If
    If Not res.QuantApplicable Then wsChart.Range("B3").Value = "Н/П (п. 4.1.2) - справочно: " & quantCell
    If Not res.QualApplicable Then wsChart.Range("B4").Value = "Н/П (п. 4.1.3)"
    If Not res.QualPass And res.QualApplicable Then wsChart.Range("B4").Font.Color = RGB(192, 0, 0)
    If res.TransientType = "Колебательный" Or res.OscViolation Then wsChart.Range("B5").Font.Color = RGB(192, 0, 0)
    If res.React15Status = "Нарушение" Then wsChart.Range("B9").Font.Color = RGB(192, 0, 0)
    If Left$(res.QualClauseVerdict, 9) = "НАРУШЕНИЕ" Then wsChart.Range("B11").Font.Color = RGB(192, 0, 0)
    If Left$(res.Verdict, 9) = "НАРУШЕНИЕ" Then
        wsChart.Range("B12").Font.Color = RGB(192, 0, 0)
    ElseIf Left$(res.Verdict, 2) = "ОК" Then
        wsChart.Range("B12").Font.Color = RGB(0, 97, 0)
    Else
        wsChart.Range("B12").Font.Color = RGB(89, 89, 89)
    End If
End Sub

Private Function FindChartEndRow(ByVal wsData As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                                 ByVal t0 As Variant, ByVal intervalSec As Double) As Long
    Dim r As Long, t0d As Double, tCur As Double, tTarget As Double
    Dim bestRow As Long, bestDiff As Double, curDiff As Double

    If Not IsDate(t0) Then
        FindChartEndRow = r2
        Exit Function
    End If
    t0d = CDbl(CDate(t0))
    tTarget = t0d + intervalSec / 86400#
    bestRow = r1
    bestDiff = 1E+99
    For r = r1 To r2
        If IsDate(wsData.Cells(r, 1).Value) Then
            tCur = CDbl(CDate(wsData.Cells(r, 1).Value))
            If tCur >= tTarget Then
                FindChartEndRow = r
                Exit Function
            End If
            curDiff = Abs(tCur - tTarget)
            If curDiff < bestDiff Then
                bestDiff = curDiff
                bestRow = r
            End If
        End If
    Next r
    FindChartEndRow = bestRow
End Function

' ==========================================================
' Станционные суммы
' ==========================================================

Private Sub BuildStationAggregates(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, _
                                   ByRef st As TSettings, ByRef outRow As Long)
    Dim cfgLast As Long, r As Long
    Dim g As TGenCfg
    Dim stations As Collection, key As Variant
    Dim done As Object

    cfgLast = LastUsedRow(wsCfg)
    Set stations = New Collection
    Set done = CreateObject("Scripting.Dictionary")

    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NX
        g = ReadGenCfgEff(wsCfg, r, st)
        If g.Enabled And g.InStationSum And Len(g.Station) > 0 Then
            If Not done.Exists(g.Station) Then
                done(g.Station) = 1
                stations.Add g.Station
            End If
        End If
NX:
    Next r

    For Each key In stations
        BuildOneStationAggregate wsRaw, wsCfg, wsSummary, CStr(key), "", st, outRow
        BuildParoprovodSubAggregates wsRaw, wsCfg, wsSummary, CStr(key), st, outRow
    Next key
End Sub

Private Sub BuildParoprovodSubAggregates(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, _
                                         ByVal stationName As String, ByRef st As TSettings, ByRef outRow As Long)
    Dim cfgLast As Long, r As Long
    Dim g As TGenCfg
    Dim paropipes As Object, p As Variant

    Set paropipes = CreateObject("Scripting.Dictionary")
    cfgLast = LastUsedRow(wsCfg)
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NX
        g = ReadGenCfgEff(wsCfg, r, st)
        If g.Enabled And g.InStationSum And StationMatch(g.Station, stationName) Then
            If Len(g.Paroprovod) > 0 Then
                If paropipes.Exists(g.Paroprovod) Then
                    paropipes(g.Paroprovod) = paropipes(g.Paroprovod) + 1
                Else
                    paropipes(g.Paroprovod) = 1
                End If
            End If
        End If
NX:
    Next r

    For Each p In paropipes.Keys
        If paropipes(p) >= 1 Then
            BuildOneStationAggregate wsRaw, wsCfg, wsSummary, stationName, CStr(p), st, outRow
        End If
    Next p
End Sub

' Суммарная мощность включённых генераторов в строке r (учитываются только работающие, > порога)
Private Function SumPowerAtRow(ByVal wsRaw As Worksheet, ByRef pCols() As Long, ByVal cnt As Long, _
                               ByVal r As Long, ByVal thresholdMW As Double) As Double
    Dim i As Long, v As Double, total As Double
    total = 0
    For i = 1 To cnt
        v = NzD(wsRaw.Cells(r, pCols(i)).Value, 0)
        If v > thresholdMW Then total = total + v
    Next i
    SumPowerAtRow = total
End Function

Private Sub BuildOneStationAggregate(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByVal wsSummary As Worksheet, _
                                      ByVal stationName As String, ByVal paropipeFilter As String, _
                                      ByRef st As TSettings, ByRef outRow As Long)
    Dim cfgLast As Long, r As Long, cnt As Long
    Dim g As TGenCfg
    Dim pCols() As Long, fnchArr() As Double, sArr() As Double, pnomArr() As Double, etArr() As String
    Dim pmaxArr() As Double, pminArr() As Double
    Dim freqCol As Long, timeCol As Long
    Dim isNeblok As Boolean
    Dim evList() As TEvent, nEv As Long, k As Long

    cfgLast = LastUsedRow(wsCfg)
    timeCol = FindHeaderCol(wsRaw, "Время")
    If timeCol = 0 Then Exit Sub

    cnt = 0
    isNeblok = False
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NextCfg
        g = ReadGenCfgEff(wsCfg, r, st)
        If g.Enabled And g.InStationSum Then
            If StationMatch(g.Station, stationName) Then
                If Len(paropipeFilter) = 0 Or StrComp(g.Paroprovod, paropipeFilter, vbTextCompare) = 0 Then
                    cnt = cnt + 1
                    ReDim Preserve pCols(1 To cnt)
                    ReDim Preserve fnchArr(1 To cnt)
                    ReDim Preserve sArr(1 To cnt)
                    ReDim Preserve pnomArr(1 To cnt)
                    ReDim Preserve etArr(1 To cnt)
                    ReDim Preserve pmaxArr(1 To cnt)
                    ReDim Preserve pminArr(1 To cnt)
                    pCols(cnt) = FindHeaderColFuzzy(wsRaw, g.PowerHeader)
                    If pCols(cnt) = 0 Then cnt = cnt - 1: GoTo NextCfg
                    fnchArr(cnt) = g.Fnch
                    sArr(cnt) = g.SPct
                    pnomArr(cnt) = g.PNom
                    etArr(cnt) = g.EquipType
                    pmaxArr(cnt) = g.PMax
                    pminArr(cnt) = g.PMin
                    If UCase$(g.Layout) = "НЕБЛОК" Then isNeblok = True
                    If freqCol = 0 Then freqCol = FindHeaderColFuzzy(wsRaw, g.FreqHeader)
                End If
            End If
        End If
NextCfg:
    Next r
    If cnt = 0 Then Exit Sub
    If freqCol = 0 Then Exit Sub

    nEv = FindEvents(wsRaw, timeCol, freqCol, st, MinArray(fnchArr), evList)
    For k = 1 To nEv
        BuildOneStationAggregateEvent wsRaw, wsSummary, stationName, paropipeFilter, st, outRow, _
                                      pCols, fnchArr, sArr, pnomArr, etArr, pmaxArr, pminArr, cnt, _
                                      freqCol, timeCol, isNeblok, evList(k), k
    Next k
End Sub

Private Sub BuildOneStationAggregateEvent(ByVal wsRaw As Worksheet, ByVal wsSummary As Worksheet, _
                                      ByVal stationName As String, ByVal paropipeFilter As String, _
                                      ByRef st As TSettings, ByRef outRow As Long, _
                                      ByRef pCols() As Long, ByRef fnchArr() As Double, ByRef sArr() As Double, _
                                      ByRef pnomArr() As Double, ByRef etArr() As String, _
                                      ByRef pmaxArr() As Double, ByRef pminArr() As Double, ByVal cnt As Long, _
                                      ByVal freqCol As Long, ByVal timeCol As Long, ByVal isNeblok As Boolean, _
                                      ByRef ev As TEvent, ByVal evIdx As Long)
    Dim r As Long
    Dim shName As String, ws As Worksheet
    Dim startRow As Long, endRow As Long, rowQ As Long, firstExceedRow As Long, evEnd As Long
    Dim p0 As Double, pNow As Double, preq As Double, preqOrig As Double
    Dim pfact As Double, dF As Double
    Dim i As Long, dfr As Double
    Dim suffix As String
    Dim pMaxSum As Double, pMinSum As Double, pNomSum As Double
    Dim reservePlus As Double, reserveMinus As Double
    Dim limited As Boolean, limitType As String
    Dim sr As TStationResult
    Dim srcLabel As String

    srcLabel = stationName & IIf(Len(paropipeFilter) > 0, "/" & paropipeFilter, "") & " (сумма)" & IIf(evIdx > 1, " событие " & evIdx, "")

    startRow = ev.StartRow
    firstExceedRow = ev.FirstExceedRow
    evEnd = ev.EndRow
    If evEnd < startRow Then evEnd = startRow
    endRow = RowByTimeOffset(wsRaw, timeCol, startRow, st.QuantIntervalSec)

    ' П. 4.2.10: сумма по включённому в работу (P > порога) и готовому к ОПРЧ оборудованию.
    ' Pном_сум / Pmax_сум / Pmin_сум - только по генераторам, находившимся в работе на старте.
    p0 = 0
    pMaxSum = 0
    pMinSum = 0
    pNomSum = 0
    Dim inWork() As Boolean
    ReDim inWork(1 To cnt)
    For i = 1 To cnt
        inWork(i) = (NzD(wsRaw.Cells(startRow, pCols(i)).Value, 0) > st.WorkThresholdMW)
        If inWork(i) Then
            p0 = p0 + NzD(wsRaw.Cells(startRow, pCols(i)).Value, 0)
            pMaxSum = pMaxSum + pmaxArr(i)
            pMinSum = pMinSum + pminArr(i)
            pNomSum = pNomSum + pnomArr(i)
        End If
    Next i
    Dim rA As Long, r1 As Long, sumA As Double, cntA As Long
    If st.AvgWindowSec > 0 Then
        ' п. 4.2.6: усреднение суммарной мощности на окне перед стартом
        r1 = RowByTimeOffset(wsRaw, timeCol, startRow, -st.AvgWindowSec)
        If r1 < 2 Then r1 = 2
        For rA = r1 To startRow
            sumA = sumA + SumPowerAtRow(wsRaw, pCols, cnt, rA, st.WorkThresholdMW)
            cntA = cntA + 1
        Next rA
        If cntA > 0 Then p0 = sumA / cntA
    End If

    ' --- Таблица по точкам окна события и дефицит (п. 4.2.14.1) ---
    If Len(paropipeFilter) > 0 Then
        suffix = "_Сумма_" & paropipeFilter
    Else
        suffix = "_Сумма"
    End If
    shName = StationSheetName(stationName, paropipeFilter, evIdx)
    Set ws = EnsureSheet(shName)
    ws.Cells.Clear

    ws.Range("A10:H10").Value = Array("Время", "Частота, Гц", "Pсум, МВт", "dPсум, МВт", "Pтреб_сум, МВт", _
                                       "dPmax_сум", "dPmin_сум", "Дефицит_сум, МВт")
    rowQ = 11
    Dim dPmaxRel As Double, dPminRel As Double
    Dim defMaxPct As Double, defSec As Double, tolSum As Double
    Dim preqStep As Double, preqCap As Double, allowed As Double, have As Double, deficit As Double, dt As Double
    Dim tSecAgg As Double, fMean As Double, isQS As Boolean
    Dim trRow As Long, trDef As Double, anyQS As Boolean, trDfrRow As Long, trDfrAbs As Double
    Dim run02 As Double, best02 As Double, dfAbs As Double
    reservePlus = pMaxSum - p0
    reserveMinus = p0 - pMinSum
    If reservePlus < 0 Then reservePlus = 0
    If reserveMinus < 0 Then reserveMinus = 0
    dPmaxRel = pMaxSum - p0
    dPminRel = pMinSum - p0
    tolSum = 0.01 * pNomSum
    defMaxPct = 0
    defSec = 0
    trRow = 0: trDef = -1E+99: anyQS = False: trDfrRow = 0: trDfrAbs = 0
    run02 = 0: best02 = 0
    dF = 0
    For r = startRow To evEnd
        ws.Cells(rowQ, 1).Value = wsRaw.Cells(r, timeCol).Value
        ws.Cells(rowQ, 2).Value = wsRaw.Cells(r, freqCol).Value

        pNow = SumPowerAtRow(wsRaw, pCols, cnt, r, st.WorkThresholdMW)
        ws.Cells(rowQ, 3).Value = pNow
        ws.Cells(rowQ, 4).Value = pNow - p0

        dfAbs = NzD(wsRaw.Cells(r, freqCol).Value, st.FNom) - st.FNom
        If Abs(dfAbs) > Abs(dF) Then dF = dfAbs
        If r < evEnd Then
            dt = SecBetween(wsRaw.Cells(r, timeCol).Value, wsRaw.Cells(r + 1, timeCol).Value)
        Else
            dt = 1
        End If
        If dt <= 0 Or dt > 60 Then dt = 1
        If Abs(dfAbs) >= QUANT_DF_HZ - 0.0000001 Then
            run02 = run02 + dt
            If run02 > best02 Then best02 = run02
        Else
            run02 = 0
        End If

        preqStep = 0
        tSecAgg = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(r, timeCol).Value)
        For i = 1 To cnt
            If inWork(i) Then
                dfr = DeadbandDeviation(dfAbs, fnchArr(i))
                preqStep = preqStep + PReqCalc(sArr(i), pnomArr(i), st.FNom, DynamicKdByTime(etArr(i), tSecAgg), dfr, st.MaxReqPct)
            End If
        Next i
        ws.Cells(rowQ, 5).Value = preqStep
        ws.Cells(rowQ, 6).Value = dPmaxRel
        ws.Cells(rowQ, 7).Value = dPminRel

        ' п. 4.2.14.1: дефицит суммарной первичной мощности по точке (допуск 1 %Pном_сум);
        ' первые 15 с (переходный участок) не оцениваются - см. реакция РЧВ (п. 4.2.11)
        deficit = 0
        preqCap = ClampReserve(preqStep, reservePlus, reserveMinus)
        allowed = Abs(preqCap) - tolSum
        If allowed > 0 And tSecAgg >= 15 Then
            have = SgnNZ(preqCap) * (pNow - p0)
            deficit = allowed - have
            If deficit < 0 Then deficit = 0
        End If
        ws.Cells(rowQ, 8).Value = deficit
        If deficit > 0 And pNomSum > 0 Then
            If 100# * deficit / pNomSum > defMaxPct Then defMaxPct = 100# * deficit / pNomSum
            defSec = defSec + dt
        End If

        ' выбор tр: худшая квазиустановившаяся точка (п. 4.2.5 / Приложение)
        If tSecAgg >= 15 And preqCap <> 0 Then
            isQS = IsQuasiSteady(wsRaw, freqCol, timeCol, r, st.FNom, fMean)
            If isQS Then
                If DeadbandDeviation(fMean - st.FNom, MinArray(fnchArr)) <> 0 Then
                    anyQS = True
                    have = SgnNZ(preqCap) * (pNow - p0)
                    If (Abs(preqCap) - tolSum) - have > trDef Then trDef = (Abs(preqCap) - tolSum) - have: trRow = r
                    If Abs(fMean - st.FNom) > trDfrAbs Then trDfrAbs = Abs(fMean - st.FNom): trDfrRow = r
                End If
            End If
        End If
        rowQ = rowQ + 1
    Next r

    ' --- Момент оценки tр и показатели в нём ---
    Dim trQS As Boolean
    If anyQS Then
        If trDef > 0 Then
            ' дефицит есть - худшая точка
        Else
            trRow = IIf(trDfrRow > 0, trDfrRow, trRow)
        End If
        trQS = True
    Else
        trRow = endRow
        trQS = False
    End If
    If trRow < startRow Then trRow = startRow
    If trRow > evEnd And anyQS Then trRow = evEnd
    tSecAgg = SecBetween(wsRaw.Cells(startRow, timeCol).Value, wsRaw.Cells(trRow, timeCol).Value)

    If st.AvgWindowSec > 0 Then
        sumA = 0: cntA = 0
        r1 = RowByTimeOffset(wsRaw, timeCol, trRow, -st.AvgWindowSec)
        If r1 < 2 Then r1 = 2
        For rA = r1 To trRow
            sumA = sumA + SumPowerAtRow(wsRaw, pCols, cnt, rA, st.WorkThresholdMW)
            cntA = cntA + 1
        Next rA
        If cntA > 0 Then pNow = sumA / cntA Else pNow = SumPowerAtRow(wsRaw, pCols, cnt, trRow, st.WorkThresholdMW)
    Else
        pNow = SumPowerAtRow(wsRaw, pCols, cnt, trRow, st.WorkThresholdMW)
    End If
    pfact = pNow - p0

    Dim dfTr As Double
    If trQS Then
        dfTr = MeanFreqAround(wsRaw, freqCol, timeCol, trRow, QS_HALF_WINDOW_SEC, st.FNom) - st.FNom
    Else
        dfTr = MaxAbsDeviationInWindow(wsRaw, freqCol, startRow, trRow, st.FNom)
    End If
    preqOrig = 0
    For i = 1 To cnt
        If inWork(i) Then
            dfr = DeadbandDeviation(dfTr, fnchArr(i))
            preqOrig = preqOrig + PReqCalc(sArr(i), pnomArr(i), st.FNom, DynamicKdByTime(etArr(i), tSecAgg), dfr, st.MaxReqPct)
        End If
    Next i
    preq = preqOrig
    limited = False
    limitType = ""
    If preq > reservePlus Then
        preq = reservePlus: limited = True: limitType = "Pmax"
    ElseIf preq < -reserveMinus Then
        preq = -reserveMinus: limited = True: limitType = "Pmin"
    End If
    If limited Then
        AppendLog "INFO", srcLabel, _
                  "Pтреб_сум ограничен " & limitType & ": было " & Format(preqOrig, "0.000") & _
                  " МВт, принято " & Format(preq, "0.000") & " МВт."
    End If

    ' Количественный критерий по сумме (как для генератора)
    Dim qpct As Double, qpass As Boolean, dfrMin As Double
    dfrMin = DeadbandDeviation(dfTr, MinArray(fnchArr))
    If dfrMin = 0 Then
        qpct = 100: qpass = True
    ElseIf Abs(preq) < 0.000001 Then
        qpct = 100: qpass = True
    ElseIf SgnNZ(pfact) <> SgnNZ(preq) Then
        qpct = 0: qpass = False
    Else
        qpct = 100# * SafeDiv(Abs(pfact), Abs(preq), 0)
        qpass = (qpct >= (100# - st.QuantTolPct))
    End If

    ' Применимость количественной оценки по сумме (п. 4.1.2)
    If st.IsEmergency Then
        sr.QuantApplicable = (best02 >= QUANT_MIN_SEC_TES - 0.5)
        If sr.QuantApplicable Then
            sr.Applicability = "колич. (п. 4.1.2: |dF| >= 0.2 Гц в течение " & Format(best02, "0") & " с)"
        Else
            sr.Applicability = "колич. оценка не применяется (п. 4.1.2: |dF| >= 0.2 Гц " & Format(best02, "0") & " с < 30 с; |dF|max = " & Format(Abs(dF), "0.000") & " Гц)"
        End If
    Else
        sr.QuantApplicable = True
        sr.Applicability = "Испытание: применимость по п. 4.1.2 не проверяется"
    End If

    ' п. 4.2.11 / 4.4.2: наличие реакции РЧВ через 15 с по суммарной мощности
    Dim baseRow As Long, row15 As Long, dp15 As Double
    Dim r15Status As String, r15Pass As Boolean, r15Note As String
    Dim preq15Sum As Double, dfr15 As Double, anyOutside As Boolean
    baseRow = IIf(firstExceedRow > 0, firstExceedRow, startRow)
    row15 = RowByTimeOffset(wsRaw, timeCol, baseRow, 15#)
    dp15 = SumPowerAtRow(wsRaw, pCols, cnt, row15, st.WorkThresholdMW) - p0
    preq15Sum = 0
    anyOutside = False
    For i = 1 To cnt
        If inWork(i) Then
            dfr15 = DeadbandDeviation(NzD(wsRaw.Cells(row15, freqCol).Value, st.FNom) - st.FNom, fnchArr(i))
            If dfr15 <> 0 Then
                anyOutside = True
                preq15Sum = preq15Sum + PReqCalc(sArr(i), pnomArr(i), st.FNom, DynamicKdByTime(etArr(i), 15#), dfr15, st.MaxReqPct)
            End If
        End If
    Next i
    preq15Sum = ClampReserve(preq15Sum, reservePlus, reserveMinus)
    Dim needSignSum As Integer
    needSignSum = SgnNZ(preqOrig)
    If needSignSum = 0 Then needSignSum = -SgnNZ(DeadbandDeviation(dF, MinArray(fnchArr)))
    ClassifyReaction15 st, dp15, pNomSum, needSignSum, reservePlus, reserveMinus, _
                       React15ThresholdMW(st, pNomSum, preq15Sum), (Not anyOutside), r15Status, r15Pass, r15Note
    If r15Status = "Нарушение" Then AppendLog "WARN", srcLabel, "п. 4.2.11: " & r15Note

    ' --- Шапка листа суммы ---
    ws.Range("A1:B1").Value = Array("Станция", stationName & IIf(evIdx > 1, " (событие " & evIdx & ")", ""))
    If Len(paropipeFilter) > 0 Then
        ws.Range("A2:B2").Value = Array("Паропровод", paropipeFilter)
    Else
        ws.Range("A2:B2").Value = Array("Режим", IIf(isNeblok, "Суммарная нагрузка неблочной части (п. 4.2.10 МУ)", "Суммарная нагрузка (блочная компоновка - справочно)"))
    End If
    ws.Range("A3:B3").Value = Array("Старт (расч.)", wsRaw.Cells(startRow, timeCol).Value)
    If firstExceedRow > 0 Then
        ws.Range("A4:B4").Value = Array("Выход за fнч", wsRaw.Cells(firstExceedRow, timeCol).Value)
    Else
        ws.Range("A4:B4").Value = Array("Выход за fнч", "")
    End If
    ws.Range("A5:B5").Value = Array("P0, МВт", p0)
    ws.Range("A6:B6").Value = Array("Pтек(tр), МВт", pNow)
    ws.Range("A7:B7").Value = Array("Pтреб(tр), МВт", preq)
    ws.Range("A8:B8").Value = Array("Pфакт(tр), МВт", pfact)
    ws.Range("A9:B9").Value = Array("tр", wsRaw.Cells(trRow, timeCol).Value)

    ws.Cells(1, 4).Resize(1, 2).Value = Array("Pmax_сум, МВт", pMaxSum)
    ws.Cells(2, 4).Resize(1, 2).Value = Array("Pmin_сум, МВт", pMinSum)
    ws.Cells(3, 4).Resize(1, 2).Value = Array("Резерв '+', МВт", reservePlus)
    ws.Cells(4, 4).Resize(1, 2).Value = Array("Резерв '-', МВт", reserveMinus)
    ws.Cells(5, 4).Resize(1, 2).Value = Array("Pтреб исх., МВт", preqOrig)
    ws.Cells(6, 4).Resize(1, 2).Value = Array("Ограничение", IIf(limited, "Да (" & limitType & ")", "нет"))
    ws.Cells(7, 4).Resize(1, 2).Value = Array("Генераторов в сумме", cnt)
    ws.Cells(8, 4).Resize(1, 2).Value = Array("Pном_сум (в работе), МВт", pNomSum)
    ws.Cells(9, 4).Resize(1, 2).Value = Array("tр, с / режим", Format(tSecAgg, "0") & " с / " & IIf(trQS, "квазиустановившийся", "переходный"))

    ' Результат по сумме станции
    sr.Station = stationName
    sr.Paroprovod = paropipeFilter
    sr.Count = cnt
    sr.EventIdx = evIdx
    sr.IsNeblok = isNeblok
    sr.StartTime = wsRaw.Cells(startRow, timeCol).Value
    If firstExceedRow > 0 Then sr.FirstExceedTime = wsRaw.Cells(firstExceedRow, timeCol).Value Else sr.FirstExceedTime = ""
    sr.TrTime = wsRaw.Cells(trRow, timeCol).Value
    sr.TrSec = tSecAgg
    sr.TrQuasiSteady = trQS
    sr.P0 = p0
    sr.PTek = pNow
    sr.Df = dF
    sr.PReq = preq
    sr.PReqOrig = preqOrig
    sr.PFact = pfact
    sr.QuantPct = qpct
    sr.QuantPass = qpass
    sr.Limited = limited
    sr.LimitType = limitType
    sr.PMaxSum = pMaxSum
    sr.PMinSum = pMinSum
    sr.ReservePlus = reservePlus
    sr.ReserveMinus = reserveMinus
    sr.PNomSum = pNomSum
    sr.React15Status = r15Status
    sr.React15Pass = r15Pass
    sr.DP15 = dp15
    sr.React15Note = r15Note
    sr.DeficitMaxPct = defMaxPct
    sr.DeficitSec = defSec
    sr.SheetName = shName
    sr.ChartSheetName = StationChartSheetName(stationName, paropipeFilter, evIdx)
    sr.Verdict = BuildStationVerdict(sr, st)

    ' Блок вердикта на листе суммы (G1:H7)
    ws.Cells(1, 7).Resize(1, 2).Value = Array("Колич. %", qpct)
    ws.Cells(2, 7).Resize(1, 2).Value = Array("Колич. статус", IIf(Not sr.QuantApplicable, "Н/П (п. 4.1.2)", IIf(limited, "Ограничен " & limitType, IIf(qpass, "ОК", "Нарушение"))))
    ws.Cells(3, 7).Resize(1, 2).Value = Array("Реакция 15 с (п. 4.2.11)", r15Status & " (dP15 = " & Format(dp15, "0.000") & " МВт)")
    ws.Cells(4, 7).Resize(1, 2).Value = Array("Дефицит (п. 4.2.14.1)", _
        IIf(defSec > 0, "до " & Format(defMaxPct, "0.00") & " %Pном_сум, " & Format(defSec, "0") & " с", "нет"))
    ws.Cells(5, 7).Resize(1, 2).Value = Array("Применимость", sr.Applicability)

    ws.Columns("A:H").AutoFit
    ws.Cells(6, 7).Value = "Вывод"
    ws.Cells(6, 8).Value = sr.Verdict
    ws.Cells(6, 7).Font.Bold = True
    If Left$(sr.Verdict, 9) = "НАРУШЕНИЕ" Then ws.Cells(6, 8).Font.Color = RGB(192, 0, 0) Else ws.Cells(6, 8).Font.Color = RGB(0, 97, 0)

    ws.Range("A1:A9").NumberFormat = "@"
    ws.Range("B1:B2").NumberFormat = "@"
    ws.Range("B3:B4").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B5:B8").NumberFormat = "0.000"
    ws.Range("B9").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("D1:D9").NumberFormat = "@"
    ws.Range("E1:E5").NumberFormat = "0.000"
    ws.Range("E6").NumberFormat = "@"
    ws.Range("E7").NumberFormat = "0"
    ws.Range("E8").NumberFormat = "0.000"
    ws.Range("E9").NumberFormat = "@"
    ws.Range("G1:G6").NumberFormat = "@"
    ws.Range("H1").NumberFormat = "0.0"
    ws.Range("H2:H6").NumberFormat = "@"
    ws.Range("A10:H10").NumberFormat = "@"
    ws.Range("A11:A100000").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("B11:H100000").NumberFormat = "0.000"

    AppendLog "INFO", srcLabel, sr.Verdict

    WriteStationChartSheet stationName, paropipeFilter, rowQ - 1, st, MinArray(fnchArr), dF, sr

    ' Строка в Summary
    WriteSummaryStationRow wsSummary, outRow, sr
    AddProtocolStation sr
    outRow = outRow + 1

    ' п. 4.2.10: пометка в строках генераторов станции о статусе суммы (только неблочная часть, первое событие)
    If Len(paropipeFilter) = 0 And isNeblok And evIdx = 1 Then
        AnnotateGeneratorRowsWithStationVerdict wsSummary, outRow - 1, stationName, (Left$(sr.Verdict, 2) = "ОК")
    End If
End Sub

Private Sub AnnotateGeneratorRowsWithStationVerdict(ByVal ws As Worksheet, ByVal lastRow As Long, _
                                                    ByVal stationName As String, ByVal stationOk As Boolean)
    Dim r As Long, addText As String, cur As String
    If stationOk Then
        addText = "Станция в целом: ОК (п. 4.2.10 МУ - индивидуальная оценка приведена справочно)"
    Else
        addText = "Станция в целом: не соответствует (п. 4.2.10 МУ - индивидуальная оценка обязательна)"
    End If
    For r = 2 To lastRow
        If StationMatch(CStr(ws.Cells(r, 1).Value), stationName) Then
            If Left$(CStr(ws.Cells(r, 2).Value), 5) <> "СУММА" Then
                cur = CStr(ws.Cells(r, 40).Value)
                If InStr(cur, "Станция в целом:") = 0 Then
                    ws.Cells(r, 40).Value = cur & IIf(Len(cur) > 0, " | ", "") & addText
                End If
            End If
        End If
    Next r
End Sub

' Вывод по сумме станции с учётом п. 4.2.10 МУ (индивидуальная оценка генераторов
' требуется только при нарушении по сумме)
Private Function BuildStationVerdict(ByRef sr As TStationResult, ByRef st As TSettings) As String
    Dim bad As String, remark As String, txt As String, objName As String
    objName = "по сумме " & IIf(Len(sr.Paroprovod) > 0, "паропровода " & sr.Paroprovod & " ", "станции ")
    bad = ""
    remark = ""

    If Not sr.QuantApplicable And sr.React15Status <> "Нарушение" Then
        txt = "НЕ ОЦЕНИВАЕТСЯ " & objName & ": " & sr.Applicability
        If Not sr.IsNeblok Then txt = txt & ". Блочная компоновка: п. 4.2.10 МУ не применяется, оценка - по каждому генератору"
        BuildStationVerdict = txt
        Exit Function
    End If

    If sr.QuantApplicable Then
        If sr.Limited Then
            remark = remark & "Pтреб_сум ограничен " & sr.LimitType & " (п. 4.2.8: оценка по располагаемому резерву); "
        ElseIf Not sr.QuantPass Then
            bad = bad & "Pфакт_сум(tр) = " & Format(sr.PFact, "0.00") & " МВт ниже Pтреб_сум(tр) = " & Format(sr.PReq, "0.00") & _
                  " МВт (" & Format(sr.QuantPct, "0") & " %) - п. 4.2.14.1; "
        End If
        If sr.DeficitSec > 0 Then
            remark = remark & "дефицит до " & Format(sr.DeficitMaxPct, "0.0") & " %Pном_сум в течение " & _
                     Format(sr.DeficitSec, "0") & " с (п. 4.2.14.1); "
        End If
    Else
        remark = remark & sr.Applicability & "; "
    End If
    If sr.React15Status = "Нарушение" Then
        bad = bad & "отсутствует реакция РЧВ турбин через 15 с (dP15_сум = " & Format(sr.DP15, "0.00") & " МВт) - п. 4.2.11 / 4.4.2; "
    End If

    If sr.IsNeblok Then
        If Len(bad) > 0 Then
            txt = "НАРУШЕНИЕ " & objName & "(неблочная часть, " & sr.Count & " ген.): " & bad & _
                  "По п. 4.2.10 МУ требуется индивидуальная оценка каждого генератора - см. строки генераторов станции. "
        Else
            txt = "ОК " & objName & "(неблочная часть, " & sr.Count & " ген.): количественно (п. 4.2.14.1) и по реакции 15 с (п. 4.2.11) соответствует. " & _
                  "По п. 4.2.10 МУ индивидуальная оценка генераторов не требуется (строки генераторов приведены справочно). "
        End If
    Else
        If Len(bad) > 0 Then
            txt = "НАРУШЕНИЕ " & objName & "(" & sr.Count & " ген., блочная компоновка - справочно): " & bad & _
                  "Нормативная оценка - по каждому генератору (п. 4.2.10 применяется к неблочной части ТЭС). "
        Else
            txt = "ОК " & objName & "(" & sr.Count & " ген., блочная компоновка - справочно): количественно и по реакции 15 с соответствует. " & _
                  "Нормативная оценка - по каждому генератору (п. 4.2.10 применяется к неблочной части ТЭС). "
        End If
    End If
    If Len(remark) > 0 Then txt = txt & "Замечания: " & remark
    txt = Trim$(txt)
    If Right$(txt, 1) = ";" Then txt = Left$(txt, Len(txt) - 1)
    BuildStationVerdict = txt
End Function

Private Function StationSheetName(ByVal stationName As String, ByVal paropipeFilter As String, ByVal evIdx As Long) As String
    Dim suffix As String, base As String
    If Len(paropipeFilter) > 0 Then
        suffix = "_Сумма_" & paropipeFilter
    Else
        suffix = "_Сумма"
    End If
    suffix = suffix & EventSuffix(evIdx)
    base = stationName
    If Len(base & suffix) > 31 Then base = Left$(base, 31 - Len(suffix))
    StationSheetName = MakeSheetName(base & suffix)
End Function

Private Function StationChartSheetName(ByVal stationName As String, ByVal paropipeFilter As String, ByVal evIdx As Long) As String
    Dim suffix As String, dataName As String, chartName As String, base As String
    If Len(paropipeFilter) > 0 Then
        suffix = "_Сумма_" & paropipeFilter
    Else
        suffix = "_Сумма"
    End If
    suffix = suffix & EventSuffix(evIdx)
    dataName = StationSheetName(stationName, paropipeFilter, evIdx)
    base = stationName
    If Len(base & suffix & CHART_SUFFIX) > 31 Then base = Left$(base, 31 - Len(suffix) - Len(CHART_SUFFIX))
    chartName = MakeSheetName(base & suffix & CHART_SUFFIX)
    If StrComp(chartName, dataName, vbTextCompare) = 0 Then
        chartName = Left$(dataName, 28) & "_Гр"
    End If
    StationChartSheetName = chartName
End Function

Private Sub WriteStationChartSheet(ByVal stationName As String, ByVal paropipeFilter As String, _
                                   ByVal lastDataRow As Long, ByRef st As TSettings, _
                                   ByVal fnchMin As Double, ByVal dFmax As Double, ByRef sr As TStationResult)
    Dim wsData As Worksheet, wsChart As Worksheet
    Dim dataName As String, chartName As String
    Dim chartObj As ChartObject
    Dim endRow As Long, startRow As Long
    Dim ySpan As Double

    dataName = sr.SheetName
    chartName = sr.ChartSheetName

    Set wsData = ThisWorkbook.Worksheets(dataName)
    Set wsChart = EnsureSheet(chartName)
    wsChart.Cells.Clear
    Do While wsChart.ChartObjects.Count > 0
        wsChart.ChartObjects(1).Delete
    Loop

    wsChart.Range("A1").Value = "График ОПРЧ (сумма): " & stationName & IIf(Len(paropipeFilter) > 0, " / " & paropipeFilter, "") & _
                                IIf(sr.EventIdx > 1, " (событие " & sr.EventIdx & ")", "")
    wsChart.Range("A1").Font.Bold = True
    wsChart.Range("A1").Font.Size = 14

    ' Блок вердикта по сумме (A3:E6)
    wsChart.Range("A3").Value = "Количественный:"
    wsChart.Range("B3").Value = IIf(sr.Limited, "Ограничен " & sr.LimitType & " (" & Format(sr.QuantPct, "0") & " % от доступного)", _
                                    IIf(sr.QuantPass, "ОК (" & Format(sr.QuantPct, "0") & " %)", "Нарушение (" & Format(sr.QuantPct, "0") & " %)"))
    wsChart.Range("D3").Value = "Pтреб / Pфакт, МВт:"
    wsChart.Range("E3").Value = Format(sr.PReq, "0.000") & " / " & Format(sr.PFact, "0.000")
    wsChart.Range("A4").Value = "Реакция 15 с:"
    wsChart.Range("B4").Value = sr.React15Status & " (dP15 = " & Format(sr.DP15, "0.000") & " МВт; п. 4.2.11)"
    wsChart.Range("D4").Value = "Дефицит:"
    wsChart.Range("E4").Value = IIf(sr.DeficitSec > 0, "до " & Format(sr.DeficitMaxPct, "0.00") & " %Pном_сум, " & _
                                Format(sr.DeficitSec, "0") & " с", "нет")
    wsChart.Range("A5").Value = "Pmax / Pmin сум, МВт:"
    wsChart.Range("B5").Value = Format(sr.PMaxSum, "0.0") & " / " & Format(sr.PMinSum, "0.0") & _
                                " (резерв +" & Format(sr.ReservePlus, "0.0") & " / -" & Format(sr.ReserveMinus, "0.0") & ")"
    wsChart.Range("D5").Value = "Генераторов:"
    wsChart.Range("E5").Value = CStr(sr.Count) & IIf(sr.IsNeblok, " (неблочная часть)", " (блочная компоновка)")
    wsChart.Range("A6").Value = "tр / применимость:"
    wsChart.Range("B6").Value = IIf(IsDate(sr.TrTime), Format(sr.TrTime, "hh:mm:ss"), "-") & " (" & Format(sr.TrSec, "0") & " с, " & _
                                IIf(sr.TrQuasiSteady, "квазиуст.", "переходный") & "); " & sr.Applicability
    wsChart.Range("A7").Value = "Вывод:"
    wsChart.Range("B7").Value = sr.Verdict
    wsChart.Range("A3:A7").Font.Bold = True
    wsChart.Range("D3:D5").Font.Bold = True
    wsChart.Range("A3:E7").NumberFormat = "@"
    If sr.Limited Then
        wsChart.Range("B3").Font.Color = RGB(156, 87, 0)
    ElseIf Not sr.QuantPass And sr.QuantApplicable Then
        wsChart.Range("B3").Font.Color = RGB(192, 0, 0)
    End If
    If sr.React15Status = "Нарушение" Then wsChart.Range("B4").Font.Color = RGB(192, 0, 0)
    If Left$(sr.Verdict, 9) = "НАРУШЕНИЕ" Then
        wsChart.Range("B7").Font.Color = RGB(192, 0, 0)
    ElseIf Left$(sr.Verdict, 2) = "ОК" Then
        wsChart.Range("B7").Font.Color = RGB(0, 97, 0)
    Else
        wsChart.Range("B7").Font.Color = RGB(89, 89, 89)
    End If

    startRow = 11
    If lastDataRow < startRow Then Exit Sub
    endRow = FindChartEndRow(wsData, startRow, lastDataRow, wsData.Cells(startRow, 1).Value, st.ChartIntervalSec)
    If endRow < startRow Then endRow = lastDataRow

    Set chartObj = wsChart.ChartObjects.Add(10, 130, 1020, 520)
    With chartObj.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Суммарный мониторинг ОПРЧ: " & stationName & IIf(Len(paropipeFilter) > 0, " / " & paropipeFilter, "") & _
                           IIf(sr.EventIdx > 1, " (событие " & sr.EventIdx & ")", "")
        On Error Resume Next
        .Axes(xlCategory).CategoryType = xlCategoryScale
        On Error GoTo 0
    End With

    AddSeries chartObj.Chart, wsData, startRow, endRow, 1, 4, "dPсум, МВт", False
    AddSeries chartObj.Chart, wsData, startRow, endRow, 1, 5, "Pтреб_сум, МВт", False
    AddLimitLineSeries chartObj.Chart, wsData, startRow, endRow, 1, 6, "Pmax_сум (dPmax)"
    AddLimitLineSeries chartObj.Chart, wsData, startRow, endRow, 1, 7, "Pmin_сум (dPmin)"
    AddSeries chartObj.Chart, wsData, startRow, endRow, 1, 2, "Частота, Гц", True

    On Error Resume Next
    chartObj.Chart.Axes(xlCategory).TickLabels.NumberFormat = "hh:mm:ss"
    With chartObj.Chart.Axes(xlValue, xlSecondary)
        ySpan = MaxD(Abs(dFmax) * 1.2, 2# * fnchMin)
        If ySpan < 0.1 Then ySpan = 0.1
        .MinimumScale = st.FNom - ySpan
        .MaximumScale = st.FNom + ySpan
    End With
    On Error GoTo 0
    ApplyChartAxesFormat chartObj.Chart, wsData, startRow, endRow, "dPсум, МВт", "f, Гц"
End Sub

' ==========================================================
' Summary / оформление / валидация / лог
' ==========================================================

Private Sub WriteSummaryRow(ByVal ws As Worksheet, ByVal r As Long, ByRef g As TGenCfg, ByRef res As TGenResult)
    Dim note As String
    Dim overshootStr As String
    Dim quantStatusStr As String, limitStr As String
    note = ""
    If res.Limited Then
        quantStatusStr = "Ограничен " & res.LimitType
        limitStr = res.LimitType
    ElseIf res.QuantPass Then
        quantStatusStr = "ОК"
        limitStr = ""
    Else
        quantStatusStr = "Нарушение"
        limitStr = ""
    End If

    If Not res.QuantPass And Not res.Limited Then note = "Колич. критерий не выполнен"
    If res.Limited Then note = Trim$(note & "; Pтреб ограничен " & res.LimitType)
    If Abs(res.P0) < 0.001 And Abs(res.PFact) < 0.001 Then note = "Нет первичного отклика; проверьте генератор/датчик"
    If Len(res.AmplitudeTag) > 0 And res.AmplitudeTag = "Слабое" Then _
        note = Trim$(note & "; Слабое возмущение (< 3 %Pном)")
    If Len(res.AmplitudeTag) > 0 And res.AmplitudeTag = "Избыточное" Then _
        note = Trim$(note & "; Возмущение > 10 %Pном (вне нормативного диапазона)")
    If Len(note) > 0 And Left$(note, 2) = "; " Then note = Mid$(note, 3)
    overshootStr = IIf(res.Overshoot, "Да", "")

    If res.React15Status = "Нарушение" Then note = Trim$(note & "; " & res.React15Note)
    If g.ForcedNoRM Then note = Trim$(note & "; без РМ: принято S = " & Format(g.SPct, "0.0") & " %, fнч = " & Format(g.Fnch, "0.000") & " Гц (режим «Авария»)")
    If Len(note) > 0 And Left$(note, 2) = "; " Then note = Mid$(note, 3)
    If Not res.QuantApplicable Then quantStatusStr = "Н/П"

    Dim qualStatusStr As String, qualT5Str As String, qualT10Str As String, qualStStr As String
    If res.QualApplicable Then
        qualStatusStr = IIf(res.QualPass, "ОК", "Нарушение")
        qualT5Str = IIf(res.QualT5Pass, "ОК", "Нарушение")
        qualT10Str = IIf(res.QualT10Pass, "ОК", "Нарушение")
        qualStStr = IIf(res.QualSteadyPass, "ОК", "Нарушение")
    Else
        qualStatusStr = "Н/П": qualT5Str = "Н/П": qualT10Str = "Н/П": qualStStr = "Н/П"
    End If

    ws.Cells(r, 1).Resize(1, 52).Value = Array( _
        g.Station, g.Generator, g.EquipType, _
        res.StartTime, res.FirstExceedTime, _
        res.P0, res.PTek, res.PsteadyAvg, _
        res.Df, res.Dfr, _
        res.PReq, res.PFact, _
        Round(res.AmplPctPnom, 2), res.AmplitudeTag, _
        res.QuantPct, quantStatusStr, overshootStr, _
        res.TransientType, res.NumExtrema, _
        qualStatusStr, qualT5Str, qualT10Str, qualStStr, _
        res.QualFailedList, _
        FormatSecOrNA(res.T5FactSec), FormatSecOrNA(res.T10FactSec), _
        GeneratorSheetName(g, res.EventIdx), GeneratorChartSheetName(g, res.EventIdx), _
        res.PMaxEff, res.PMinEff, res.ReservePlus, res.ReserveMinus, _
        res.PReqOrig, limitStr, _
        note, _
        res.React15Status, Round(res.DP15, 3), Round(res.DeficitMaxPct, 2), Round(res.DeficitSec, 0), _
        res.Verdict, _
        res.EventIdx, res.Applicability, res.TrTime, Round(res.TrSec, 0), _
        IIf(res.TrQuasiSteady, "Квазиуст.", "Переходный"), _
        Round(res.PReqRef, 3), Round(res.PFactRef, 3), _
        IIf(res.OscViolation, "Нарушение: " & res.OscNote, IIf(Len(res.OscNote) > 0, res.OscNote, "нет")), _
        IIf(Len(res.TelemetryNote) > 0, res.TelemetryNote, "ОК"), _
        IIf(Len(res.ReactionType) > 0, res.ReactionType, "-"), _
        IIf(res.H15 > -900, Round(res.H15, 3), ""), _
        IIf(Len(res.QualClauseVerdict) > 0, res.QualClauseVerdict, IIf(res.QualApplicable, "", "Н/П (п. 4.1.3)")) _
    )
End Sub

Private Sub WriteSummaryInvalid(ByVal ws As Worksheet, ByVal r As Long, ByRef g As TGenCfg)
    ws.Cells(r, 1).Resize(1, 52).Value = Array( _
        g.Station, g.Generator, g.EquipType, "", "", "", "", "", "", "", "", "", "", "", "", _
        "Нарушение", "", "", 0, "Н/Д", "Н/Д", "Н/Д", "Н/Д", "", "", "", "Config", "", _
        "", "", "", "", "", "", _
        "Не заполнен обязательный параметр config", _
        "Н/Д", "", "", "", "НАРУШЕНИЕ: не заполнены обязательные параметры Config", _
        "", "", "", "", "", "", "", "", "", "", "", "" _
    )
End Sub

' Строка Summary по сумме станции (п. 4.2.10)
Private Sub WriteSummaryStationRow(ByVal ws As Worksheet, ByVal r As Long, ByRef sr As TStationResult)
    Dim quantStatusStr As String, limitStr As String, genLabel As String
    If sr.Limited Then
        quantStatusStr = "Ограничен " & sr.LimitType
        limitStr = sr.LimitType
    ElseIf sr.QuantPass Then
        quantStatusStr = "ОК"
        limitStr = ""
    Else
        quantStatusStr = "Нарушение"
        limitStr = ""
    End If
    genLabel = "СУММА" & IIf(Len(sr.Paroprovod) > 0, " (" & sr.Paroprovod & ")", "")
    Dim amplVal As Variant
    If sr.PNomSum > 0 Then amplVal = Round(100# * Abs(sr.PReqOrig) / sr.PNomSum, 2) Else amplVal = ""

    If Not sr.QuantApplicable Then quantStatusStr = "Н/П"

    ws.Cells(r, 1).Resize(1, 52).Value = Array( _
        sr.Station, genLabel, "Сумма станции (" & sr.Count & " ген., " & IIf(sr.IsNeblok, "неблок", "блок") & ")", _
        sr.StartTime, sr.FirstExceedTime, _
        sr.P0, sr.PTek, "", _
        sr.Df, "", _
        sr.PReq, sr.PFact, _
        amplVal, "", _
        sr.QuantPct, quantStatusStr, "", _
        "", "", _
        "Н/П", "Н/П", "Н/П", "Н/П", _
        "", "", "", _
        sr.SheetName, sr.ChartSheetName, _
        sr.PMaxSum, sr.PMinSum, sr.ReservePlus, sr.ReserveMinus, _
        sr.PReqOrig, limitStr, _
        IIf(sr.React15Status = "Нарушение", sr.React15Note, IIf(sr.IsNeblok, "Оценка по суммарной мощности неблочной части (п. 4.2.10 МУ)", "Сумма по блочной станции - справочно")), _
        sr.React15Status, Round(sr.DP15, 3), Round(sr.DeficitMaxPct, 2), Round(sr.DeficitSec, 0), _
        sr.Verdict, _
        sr.EventIdx, sr.Applicability, sr.TrTime, Round(sr.TrSec, 0), IIf(sr.TrQuasiSteady, "Квазиуст.", "Переходный"), _
        "", "", "", "", "", "", "" _
    )
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 52)).Font.Bold = True
End Sub

Private Sub ApplySummaryConditionalFormat(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim rng As Range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    ws.Range("A2:AZ" & lastRow).Interior.Pattern = xlNone

    ' Строки сумм станций (B начинается с "СУММА") - подсветка A:C
    Dim rr As Long
    For rr = 2 To lastRow
        If Left$(CStr(ws.Cells(rr, 2).Value), 5) = "СУММА" Then
            ws.Range(ws.Cells(rr, 1), ws.Cells(rr, 3)).Interior.Color = RGB(221, 235, 247)
        End If
    Next rr

    ' Колонки текстовых статусов: P, Q, T, U, V, W (колич/перерег/кач/t5/t10/уст)
    ApplyStatusCF ws, ws.Range("P2:P" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("T2:T" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("U2:U" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("V2:V" & lastRow), "ОК", "Нарушение"
    ' Реакция 15 с (AJ) и Вывод (AN)
    ApplyStatusCF ws, ws.Range("AJ2:AJ" & lastRow), "ОК", "Нарушение"
    ApplyStatusCF ws, ws.Range("AN2:AN" & lastRow), "ОК", "НАРУШЕНИЕ"
    ' Колебания (AV), Кач. вывод (AZ): "Нарушение" красным; Качество ТИ (AW): "сомнительны" оранжевым;
    ' Тип реакции (AX): «адекватная» зелёным, прочие оранжевым; Применимость (AP): "не удовлетворяет"/"не применяется" серым
    ApplyStatusCF ws, ws.Range("AV2:AV" & lastRow), "нет", "Нарушение"
    ApplyStatusCF ws, ws.Range("AZ2:AZ" & lastRow), "выполнены", "НАРУШЕНИЕ"
    Set rng = ws.Range("AW2:AW" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="сомнительны", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    Set rng = ws.Range("AX2:AX" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="адекватная", TextOperator:=xlContains)
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:=" - «", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    Set rng = ws.Range("AP2:AP" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="не удовлетворяет", TextOperator:=xlContains)
        .Interior.Color = RGB(217, 217, 217)
        .Font.Color = RGB(89, 89, 89)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Только", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
    End With
    ' Дефицит (AL): любое значение > 0 - оранжевым (замечание)
    Set rng = ws.Range("AL2:AL" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    
    ' Кач.уст (W): "Нарушение" подсвечиваем как замечание (оранжевым), не красным.
    Set rng = ws.Range("W2:W" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Нарушение", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="ОК", TextOperator:=xlContains)
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With

    ' Количественный статус P: 'Ограничен' -> жёлтым
    Set rng = ws.Range("P2:P" & lastRow)
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Ограничен", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With

    ' Ограничение AH (34): любой непустой текст -> жёлтым
    Set rng = ws.Range("AH2:AH" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Pmax", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Pmin", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With

    ' Перерегулирование (Q) - желтый при "Да"
    Set rng = ws.Range("Q2:Q" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Да", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 87, 0)
    End With

    ' Характер процесса (R)
    Set rng = ws.Range("R2:R" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Колебательный", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Апериодический", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Монотонный", TextOperator:=xlContains)
        .Interior.Color = RGB(198, 239, 206)
    End With

    ' Масштаб события (N)
    Set rng = ws.Range("N2:N" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Слабое", TextOperator:=xlContains)
        .Interior.Color = RGB(217, 225, 242)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Избыточное", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 217, 102)
    End With

    ' Колич. % (O) - цветовая шкала 0..100..200
    Set rng = ws.Range("O2:O" & lastRow)
    rng.FormatConditions.Delete
    With rng.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).Type = xlConditionValueNumber
        .ColorScaleCriteria(1).Value = 0
        .ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
        .ColorScaleCriteria(2).Type = xlConditionValueNumber
        .ColorScaleCriteria(2).Value = 100
        .ColorScaleCriteria(2).FormatColor.Color = RGB(99, 190, 123)
        .ColorScaleCriteria(3).Type = xlConditionValueNumber
        .ColorScaleCriteria(3).Value = 200
        .ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)
    End With

    ' Вывод (AN): "НЕ ОЦЕНИВАЕТСЯ" серым
    Set rng = ws.Range("AN2:AN" & lastRow)
    With rng.FormatConditions.Add(Type:=xlTextString, String:="НЕ ОЦЕНИВАЕТСЯ", TextOperator:=xlContains)
        .Interior.Color = RGB(217, 217, 217)
        .Font.Color = RGB(89, 89, 89)
    End With
    ' Колич. статус (P): "Н/П" серым
    Set rng = ws.Range("P2:P" & lastRow)
    With rng.FormatConditions.Add(Type:=xlTextString, String:="Н/П", TextOperator:=xlContains)
        .Interior.Color = RGB(217, 217, 217)
        .Font.Color = RGB(89, 89, 89)
    End With

    ws.Range("A1:AZ1").Font.Bold = True
    ws.Range("A1:AZ1").Interior.Color = RGB(217, 225, 242)
End Sub

Private Sub ApplyStatusCF(ByVal ws As Worksheet, ByVal rng As Range, ByVal okText As String, ByVal badText As String)
    rng.FormatConditions.Delete
    With rng.FormatConditions.Add(Type:=xlTextString, String:=badText, TextOperator:=xlContains)
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With
    With rng.FormatConditions.Add(Type:=xlTextString, String:=okText, TextOperator:=xlContains)
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With
End Sub

Private Sub WriteVersionStamp(ByVal wsSummary As Worksheet, ByVal wsRaw As Worksheet, ByVal t0Run As Double)
    Dim col As Long
    col = 54   ' колонка BB - правее данных Summary (после AZ)
    wsSummary.Cells(1, col).Value = "Мониторинг ОПРЧ"
    wsSummary.Cells(2, col).Value = "Версия"
    wsSummary.Cells(2, col + 1).Value = OPRCH_VERSION
    wsSummary.Cells(3, col).Value = "Запуск"
    wsSummary.Cells(3, col + 1).Value = Now
    wsSummary.Cells(3, col + 1).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsSummary.Cells(4, col).Value = "Длительность, с"
    wsSummary.Cells(4, col + 1).Value = Round(Timer - t0Run, 1)
    wsSummary.Cells(5, col).Value = "Файл книги"
    wsSummary.Cells(5, col + 1).Value = ThisWorkbook.Name
    wsSummary.Cells(6, col).Value = "Строк RawData"
    wsSummary.Cells(6, col + 1).Value = LastUsedRow(wsRaw) - 1
    wsSummary.Range(wsSummary.Cells(1, col), wsSummary.Cells(6, col)).Font.Bold = True
    wsSummary.Range(wsSummary.Cells(1, col), wsSummary.Cells(6, col + 1)).Interior.Color = RGB(240, 240, 240)
End Sub

' ==========================================================
' Валидация и лог
' ==========================================================

Private Sub InitLog()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SH_LOG)
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Уровень", "Источник", "Сообщение", "Время")
    ws.Range("A1:D1").Font.Bold = True
    m_LogRow = 2
End Sub

Private Sub AppendLog(ByVal level As String, ByVal source As String, ByVal message As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0
    If ws Is Nothing Then InitLog: Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If m_LogRow < 2 Then m_LogRow = 2
    ws.Cells(m_LogRow, 1).Value = level
    ws.Cells(m_LogRow, 2).Value = source
    ws.Cells(m_LogRow, 3).Value = message
    ws.Cells(m_LogRow, 4).Value = Now
    ws.Cells(m_LogRow, 4).NumberFormat = "hh:mm:ss"
    m_LogRow = m_LogRow + 1
End Sub

Private Sub FinalizeLog()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ws.Columns("A:D").AutoFit
    ws.Range("A1:D1").Interior.Color = RGB(217, 225, 242)
    If m_LogRow <= 2 Then
        ws.Range("A2").Value = "INFO"
        ws.Range("B2").Value = "-"
        ws.Range("C2").Value = "Замечаний не обнаружено"
        ws.Range("D2").Value = Now
        ws.Range("D2").NumberFormat = "hh:mm:ss"
    End If
End Sub

Private Sub ValidateInputs(ByVal wsRaw As Worksheet, ByVal wsCfg As Worksheet, ByRef st As TSettings, ByVal timeCol As Long)
    Dim lastR As Long, r As Long, prev As Double, cur As Double, gap As Double, maxGap As Double
    Dim cfgLast As Long, g As TGenCfg, headerRow As Long
    Dim paropipes As Object, stationParopipe As Object, st2 As String, kk As Variant

    ' RawData: шаг
    lastR = LastUsedRow(wsRaw)
    If lastR < 10 Then
        AppendLog "WARN", "RawData", "Слишком мало строк данных: " & (lastR - 1)
        Exit Sub
    End If
    prev = 0
    maxGap = 0
    For r = 2 To WorksheetFunction.Min(lastR, 500)
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then
            cur = CDbl(CDate(wsRaw.Cells(r, timeCol).Value))
            If prev > 0 Then
                gap = (cur - prev) * 86400#
                If gap > maxGap Then maxGap = gap
                If gap > 6 Then AppendLog "WARN", "RawData", "Разрыв по времени " & Format(gap, "0.0") & " c в строке " & r & " (норматив <= 5 c)"
            End If
            prev = cur
        End If
    Next r
    If maxGap > 0 Then AppendLog "INFO", "RawData", "Максимальный шаг по времени в первых строках: " & Format(maxGap, "0.0") & " c"

    ' Config: валидация параметров
    cfgLast = LastUsedRow(wsCfg)
    Set stationParopipe = CreateObject("Scripting.Dictionary")
    For r = 2 To cfgLast
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then GoTo NX
        g = ReadGenCfgEff(wsCfg, r, st)
        If Not g.Enabled Then GoTo NX
        If g.ForcedNoRM Then AppendLog "INFO", g.Station & "/" & g.Generator, _
            "Без РМ, режим «Авария»: принято S = " & NO_RM_S_PCT & " %, fнч = " & NO_RM_FNCH_HZ & " Гц (в Config S = " & g.SPctCfg & " %, fнч = " & g.FnchCfg & " Гц)"
        If Not g.HasRM And Not st.IsEmergency Then AppendLog "INFO", g.Station & "/" & g.Generator, _
            "Без РМ, режим «Испытание»: используются фактические S = " & g.SPct & " %, fнч = " & g.Fnch & " Гц из Config"
        If g.PNom <= 0 Then AppendLog "WARN", g.Station & "/" & g.Generator, "Pном <= 0"
        If g.SPct <= 1 Or g.SPct > 15 Then AppendLog "WARN", g.Station & "/" & g.Generator, "S вне [1..15] %: " & g.SPct
        If g.Fnch < 0 Or g.Fnch > 0.5 Then AppendLog "WARN", g.Station & "/" & g.Generator, "fнч вне [0..0.5] Гц: " & g.Fnch
        If g.T10Sec < g.T5Sec Then AppendLog "WARN", g.Station & "/" & g.Generator, "t10 < t5 (" & g.T10Sec & " < " & g.T5Sec & ")"
        ' Pmax/Pmin
        If g.PMax > 0 And g.PMax < g.PMin Then _
            AppendLog "WARN", g.Station & "/" & g.Generator, "Pmax < Pmin (" & g.PMax & " < " & g.PMin & ")"
        If g.PMax > 0 And g.PNom > 0 And g.PMax > 1.3 * g.PNom Then _
            AppendLog "WARN", g.Station & "/" & g.Generator, "Pmax > 1.3*Pном (проверьте ед. измерения)"
        If g.PMin < 0 Then _
            AppendLog "WARN", g.Station & "/" & g.Generator, "Pmin < 0: " & g.PMin
        ' Паропровод: все или никто в рамках станции
        If g.InStationSum Then
            If stationParopipe.Exists(g.Station) Then
                st2 = CStr(stationParopipe(g.Station))
                If (st2 = "YES" And Len(g.Paroprovod) = 0) Or (st2 = "NO" And Len(g.Paroprovod) > 0) Then
                    AppendLog "WARN", g.Station, "Колонка Паропровод заполнена частично - проверьте согласованность"
                End If
            Else
                stationParopipe(g.Station) = IIf(Len(g.Paroprovod) > 0, "YES", "NO")
            End If
        End If
NX:
    Next r

    If st.QuantIntervalSec < 60 Then AppendLog "WARN", "Settings", "Колич. интервал < 60 с: " & st.QuantIntervalSec
    If st.SteadyWindowSec < 10 Then AppendLog "WARN", "Settings", "Окно установившегося < 10 с: " & st.SteadyWindowSec
    If st.React15ThrPct > 2 Then AppendLog "WARN", "Settings", "Порог реакции 15 с > 2 %Pном (" & st.React15ThrPct & ") - п. 4.2.11 требует лишь наличия реакции"
    If st.AvgWindowSec > 30 Then AppendLog "WARN", "Settings", "Окно усреднения P0/Pтек > 30 с: " & st.AvgWindowSec
    If st.MaxReqPct > 0 And st.MaxReqPct < 10 Then AppendLog "WARN", "Settings", "Ограничение Pтреб < 10 %Pном (" & st.MaxReqPct & ") - по Приложению к МУ принимается 10 %"
    If st.MaxEventSec < st.QuantIntervalSec Then AppendLog "WARN", "Settings", "Макс. окно события (" & st.MaxEventSec & " с) меньше контрольного интервала (" & st.QuantIntervalSec & " с)"

    ' Справка по профилям Kд(t): значения в контрольных точках 15 с и колич. интервала
    If Not m_KdProfiles Is Nothing Then
        For Each kk In m_KdProfiles.Keys
            On Error Resume Next
            AppendLog "INFO", "Kд(t) " & CStr(kk), "Kд(15 с) = " & Format(DynamicKdByTime(CStr(kk), 15#), "0.00") & _
                      "; Kд(" & Format(st.QuantIntervalSec, "0") & " с) = " & Format(DynamicKdByTime(CStr(kk), st.QuantIntervalSec), "0.00") & _
                      " (МУ: ТЭС 0.50 до 15 с -> 1.0 к 300/360/420 с; ПГУ 0.25 до 15 с -> 0.5 к 30 с -> 1.0 к 120 с)"
            On Error GoTo 0
        Next kk
    End If
End Sub

' ==========================================================
' Параметры / пресеты / чтение конфига
' ==========================================================

Private Function ReadSettings(ByVal wsCfg As Worksheet) As TSettings
    Dim st As TSettings
    Dim valCol As Long, modeTxt As String
    ' Глобальные настройки - в колонках W/X (23/24). Совместимость с книгами
    ' до 1.3.x (настройки в T/U) снята в 1.7.0: колонки U/V заняты признаками
    ' «РМ» и «Компоновка». Для старых книг запустите SetupOPRCHTemplate.
    valCol = 24
    If Trim$(CStr(wsCfg.Cells(2, 23).Value)) = "" Then
        Err.Raise vbObjectError + 2003, , "В Config не найдены глобальные настройки (W2:X16). Запустите SetupOPRCHTemplate или заполните колонки W/X."
    End If
    st.FNom = NzD(wsCfg.Cells(2, valCol).Value, 50#)
    st.AutoStart = (NzD(wsCfg.Cells(4, valCol).Value, 1) <> 0)
    st.QuantIntervalSec = NzD(wsCfg.Cells(5, valCol).Value, 82#)
    st.QuantTolPct = NzD(wsCfg.Cells(6, valCol).Value, 10#)
    st.WorkThresholdMW = NzD(wsCfg.Cells(7, valCol).Value, 1#)
    st.PreBufferSec = NzD(wsCfg.Cells(8, valCol).Value, 5#)
    st.ChartIntervalSec = NzD(wsCfg.Cells(9, valCol).Value, 120#)
    st.SteadyWindowSec = NzD(wsCfg.Cells(10, valCol).Value, 30#)
    st.React15ThrPct = NzD(wsCfg.Cells(11, valCol).Value, 0.5)
    st.AvgWindowSec = NzD(wsCfg.Cells(12, valCol).Value, 0#)
    If st.React15ThrPct < 0 Then st.React15ThrPct = 0
    If st.AvgWindowSec < 0 Then st.AvgWindowSec = 0
    ' 1.7.0: режим, окно события, число событий, ограничение Pтреб
    modeTxt = UCase$(Trim$(CStr(wsCfg.Cells(13, valCol).Value)))
    If Len(modeTxt) = 0 Then modeTxt = "АВАРИЯ"
    If Left$(modeTxt, 4) = "ИСПЫ" Or Left$(modeTxt, 4) = "ТЕСТ" Or Left$(modeTxt, 4) = "TEST" Then
        st.Mode = "Испытание"
        st.IsEmergency = False
    Else
        st.Mode = "Авария"
        st.IsEmergency = True
    End If
    st.MaxEventSec = NzD(wsCfg.Cells(14, valCol).Value, 600#)
    If st.MaxEventSec < 60 Then st.MaxEventSec = 60
    st.MaxEvents = CLng(NzD(wsCfg.Cells(15, valCol).Value, 3#))
    If st.MaxEvents < 1 Then st.MaxEvents = 1
    st.MaxReqPct = NzD(wsCfg.Cells(16, valCol).Value, 10#)
    If st.MaxReqPct < 0 Then st.MaxReqPct = 0
    If IsDate(wsCfg.Cells(3, valCol).Value) Then
        st.EventStart = CDate(wsCfg.Cells(3, valCol).Value)
    Else
        st.EventStart = 0
    End If
    ReadSettings = st
End Function

Private Function ReadGenCfg(ByVal ws As Worksheet, ByVal r As Long) As TGenCfg
    Dim g As TGenCfg, pr As TGenCfg
    g.Station = Trim$(CStr(ws.Cells(r, 1).Value))
    g.Generator = Trim$(CStr(ws.Cells(r, 2).Value))
    g.PowerHeader = Trim$(CStr(ws.Cells(r, 3).Value))
    g.FreqHeader = Trim$(CStr(ws.Cells(r, 4).Value))
    g.EquipType = Trim$(CStr(ws.Cells(r, 5).Value))
    g.PNom = NzD(ws.Cells(r, 6).Value, 0)
    g.SPct = NzD(ws.Cells(r, 7).Value, 0)
    g.Fnch = NzD(ws.Cells(r, 8).Value, -1)
    g.Enabled = (NzD(ws.Cells(r, 9).Value, 1) <> 0)
    g.QualEnabled = (NzD(ws.Cells(r, 10).Value, 1) <> 0)
    g.T5Sec = NzD(ws.Cells(r, 11).Value, 0)
    g.Dp5Pct = NzD(ws.Cells(r, 12).Value, 0)
    g.T10Sec = NzD(ws.Cells(r, 13).Value, 0)
    g.Dp10Pct = NzD(ws.Cells(r, 14).Value, 0)
    g.SteadyTolPct = NzD(ws.Cells(r, 15).Value, 0)
    g.InStationSum = (NzD(ws.Cells(r, 16).Value, 0) <> 0)
    g.CheckSteady = (NzD(ws.Cells(r, 17).Value, 1) <> 0)
    g.Paroprovod = Trim$(CStr(ws.Cells(r, 18).Value))
    ' Pmax/Pmin: пусто = значения по умолчанию (Pmax=Pном, Pmin=0)
    If Trim$(CStr(ws.Cells(r, 19).Value)) = "" Then
        g.PMax = g.PNom
    Else
        g.PMax = NzD(ws.Cells(r, 19).Value, g.PNom)
    End If
    If Trim$(CStr(ws.Cells(r, 20).Value)) = "" Then
        g.PMin = 0
    Else
        g.PMin = NzD(ws.Cells(r, 20).Value, 0)
    End If
    If g.PMax <= 0 Then g.PMax = g.PNom
    If g.PMin < 0 Then g.PMin = 0
    If g.PMax < g.PMin Then g.PMax = g.PMin
    ' РМ (U): пусто = 1 (оснащён); Компоновка (V): пусто = по типу оборудования
    g.HasRM = (NzD(ws.Cells(r, 21).Value, 1) <> 0)
    g.Layout = LCase$(Trim$(CStr(ws.Cells(r, 22).Value)))
    If Len(g.Layout) = 0 Then
        If InStr(UCase$(g.EquipType), "НЕБЛОК") > 0 Then g.Layout = "неблок" Else g.Layout = "блок"
    ElseIf InStr(g.Layout, "неблок") > 0 Or InStr(g.Layout, "общ") > 0 Or InStr(g.Layout, "попереч") > 0 Then
        g.Layout = "неблок"
    Else
        g.Layout = "блок"
    End If
    g.SPctCfg = g.SPct
    g.FnchCfg = g.Fnch
    g.ForcedNoRM = False

    pr = GetPreset(g.EquipType)
    If g.T5Sec <= 0 Then g.T5Sec = pr.T5Sec
    If g.Dp5Pct <= 0 Then g.Dp5Pct = pr.Dp5Pct
    If g.T10Sec <= 0 Then g.T10Sec = pr.T10Sec
    If g.Dp10Pct <= 0 Then g.Dp10Pct = pr.Dp10Pct
    If g.SteadyTolPct <= 0 Then g.SteadyTolPct = pr.SteadyTolPct
    If g.Fnch < 0 Then g.Fnch = pr.Fnch

    If Len(g.PowerHeader) = 0 Then g.PowerHeader = g.Generator
    If Len(g.FreqHeader) = 0 Then g.FreqHeader = "Частота"
    ReadGenCfg = g
End Function

Private Function GetPreset(ByVal equipType As String) As TGenCfg
    Dim g As TGenCfg, et As String
    et = UCase$(Trim$(equipType))
    ' Нормативные ступени динамики по МУ проверки готовности (Приложение 1 к
    ' Техтребованиям, табл. Е.1-Е.6, бланки-графики Ж.1-Ж.3) и параметрам Kд(t)
    ' Приложения 14:
    '   ПСУ блочные: 5 %Pном за 15 с, 10 %Pном за 5 (газомазутные) / 6 (пылеугольные) мин;
    '   ПСУ неблочные (общий паропровод): 5 % за 15 с, 10 % за 7 мин (по сумме ТГ);
    '   ГТУ / ПГУ утилизационные: 2.5 % за 15 с, 5 % за 30 с, 10 % за 2 мин;
    '   ПГУ сбросные: как ПСУ (5 % за 15 с, 10 % за 5-6 мин);
    '   ГПА: 10 % за 1 мин.
    ' Установившееся отклонение: не более ±1 %Pном (п. 6.2.2.11 Приложения 1, п. 4.2.7 Приложения 14).
    g.T5Sec = 15
    g.Dp5Pct = 5
    g.T10Sec = 420
    g.Dp10Pct = 10
    g.SteadyTolPct = 1
    g.Fnch = 0.075

    If InStr(et, "ГПА") > 0 Then
        g.T10Sec = 60
    ElseIf InStr(et, "ГТУ") > 0 Then
        g.Dp5Pct = 2.5
        g.T10Sec = 120
    ElseIf InStr(et, "ПГУ_СБРОСН") > 0 Or InStr(et, "СБРОСН") > 0 Then
        g.T10Sec = 300
    ElseIf InStr(et, "ПГУ_УТИЛ") > 0 Or InStr(et, "УТИЛИЗ") > 0 Or InStr(et, "ПГУ") > 0 Then
        g.Dp5Pct = 2.5
        g.T10Sec = 120
    ElseIf InStr(et, "УГОЛЬ") > 0 Or InStr(et, "ПЫЛЕУГ") > 0 Then
        g.T10Sec = 360
    ElseIf InStr(et, "ПТУ_БЛОК") > 0 Then
        g.T10Sec = 300
    ElseIf InStr(et, "ПТУ_НЕБЛОК") > 0 Or InStr(et, "НЕБЛОК") > 0 Then
        g.T10Sec = 420
    ElseIf InStr(et, "ПТУ") > 0 Or InStr(et, "ПСУ") > 0 Then
        ' короткое имя «ПТУ»/«ПСУ» без суффикса: ориентир блочной компоновки (5 мин)
        g.T10Sec = 300
    End If
    GetPreset = g
End Function

Private Function ValidateGenCfg(ByRef g As TGenCfg) As Boolean
    ValidateGenCfg = (Len(g.Station) > 0 And Len(g.Generator) > 0 And Len(g.PowerHeader) > 0 And Len(g.FreqHeader) > 0 _
                      And g.PNom > 0 And g.SPct > 0 And g.Fnch >= 0)
End Function

' ==========================================================
' Алгоритмические помощники
' ==========================================================

Private Function RowByTimeOffset(ByVal wsRaw As Worksheet, ByVal timeCol As Long, ByVal startRow As Long, ByVal offsetSec As Double) As Long
    Dim t0 As Double, tTarget As Double
    Dim r As Long, lastR As Long, bestRow As Long
    Dim curT As Double, bestDiff As Double, curDiff As Double
    Dim scanFrom As Long, scanTo As Long

    If Not IsDate(wsRaw.Cells(startRow, timeCol).Value) Then RowByTimeOffset = startRow: Exit Function
    t0 = CDbl(CDate(wsRaw.Cells(startRow, timeCol).Value))
    tTarget = t0 + offsetSec / 86400#
    lastR = LastUsedRow(wsRaw)
    bestRow = startRow: bestDiff = 1E+99

    If offsetSec >= 0 Then scanFrom = startRow: scanTo = lastR Else scanFrom = 2: scanTo = startRow

    For r = scanFrom To scanTo
        If IsDate(wsRaw.Cells(r, timeCol).Value) Then
            curT = CDbl(CDate(wsRaw.Cells(r, timeCol).Value))
            curDiff = Abs(CDbl(curT - tTarget))
            If curDiff < bestDiff Then bestDiff = curDiff: bestRow = r
        End If
    Next r
    RowByTimeOffset = bestRow
End Function

Private Function MaxAbsDeviationInWindow(ByVal wsRaw As Worksheet, ByVal freqCol As Long, ByVal r1 As Long, ByVal r2 As Long, ByVal fNom As Double) As Double
    Dim r As Long, d As Double, best As Double
    best = 0
    For r = r1 To r2
        d = NzD(wsRaw.Cells(r, freqCol).Value, fNom) - fNom
        If Abs(d) > Abs(best) Then best = d
    Next r
    MaxAbsDeviationInWindow = best
End Function

Private Function DeadbandDeviation(ByVal dF As Double, ByVal fnch As Double) As Double
    If Abs(dF) <= fnch Then DeadbandDeviation = 0 Else DeadbandDeviation = Sgn(dF) * (Abs(dF) - fnch)
End Function

Private Function SecBetween(ByVal t1 As Variant, ByVal t2 As Variant) As Double
    If IsDate(t1) And IsDate(t2) Then
        SecBetween = Abs(CDbl(CDate(t2) - CDate(t1))) * 86400#
    Else
        SecBetween = 0
    End If
End Function

Private Sub AddSeries(ByVal ch As Chart, ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                      ByVal xCol As Long, ByVal yCol As Long, ByVal titleText As String, ByVal secondaryAxis As Boolean)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = titleText
    s.XValues = ws.Range(ws.Cells(r1, xCol), ws.Cells(r2, xCol))
    s.Values = ws.Range(ws.Cells(r1, yCol), ws.Cells(r2, yCol))
    If secondaryAxis Then s.AxisGroup = xlSecondary
End Sub

Private Sub AddMarkerSeries(ByVal ch As Chart, ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                            ByVal xCol As Long, ByVal yCol As Long, ByVal nm As String)
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.XValues = ws.Range(ws.Cells(r1, xCol), ws.Cells(r2, xCol))
    s.Values = ws.Range(ws.Cells(r1, yCol), ws.Cells(r2, yCol))
    On Error Resume Next
    s.ChartType = xlLine
    s.Format.Line.Weight = 1.25
    s.Format.Line.DashStyle = msoLineDash
    s.MarkerStyle = xlMarkerStyleNone
    On Error GoTo 0
End Sub

Private Sub AddLimitLineSeries(ByVal ch As Chart, ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                               ByVal xCol As Long, ByVal yCol As Long, ByVal nm As String)
    ' Горизонтальные уровни Pmax / Pmin: красная жирная пунктирная линия без маркеров.
    Dim s As Series
    Set s = ch.SeriesCollection.NewSeries
    s.Name = nm
    s.XValues = ws.Range(ws.Cells(r1, xCol), ws.Cells(r2, xCol))
    s.Values = ws.Range(ws.Cells(r1, yCol), ws.Cells(r2, yCol))
    On Error Resume Next
    s.ChartType = xlLine
    s.Format.Line.Weight = 1.75
    s.Format.Line.DashStyle = msoLineLongDash
    s.Format.Line.ForeColor.RGB = RGB(192, 0, 0)
    s.MarkerStyle = xlMarkerStyleNone
    On Error GoTo 0
End Sub

Private Function AmplitudeTag(ByVal amplPct As Double, ByVal isEvent As Boolean) As String
    If Not isEvent Then AmplitudeTag = "": Exit Function
    If amplPct < 3# Then
        AmplitudeTag = "Слабое"
    ElseIf amplPct > 10# Then
        AmplitudeTag = "Избыточное"
    Else
        AmplitudeTag = "Норма"
    End If
End Function

Private Sub LoadKdProfiles(ByVal wsCfg As Worksheet)
    Dim r As Long, lastR As Long
    Dim et As String
    Dim t0 As Double, t1 As Double, t2 As Double, t3 As Double
    Dim kd0 As Double, kd1 As Double, kd2 As Double, kd3 As Double
    Set m_KdProfiles = CreateObject("Scripting.Dictionary")
    lastR = LastUsedRow(wsCfg)
    For r = 3 To lastR
        et = UCase$(Trim$(CStr(wsCfg.Cells(r, 27).Value)))
        If Len(et) = 0 Then GoTo NX
        t0 = NzD(wsCfg.Cells(r, 28).Value, 0)
        kd0 = NzD(wsCfg.Cells(r, 29).Value, -1)
        t1 = NzD(wsCfg.Cells(r, 30).Value, 0)
        kd1 = NzD(wsCfg.Cells(r, 31).Value, -1)
        t2 = NzD(wsCfg.Cells(r, 32).Value, 0)
        kd2 = NzD(wsCfg.Cells(r, 33).Value, -1)
        t3 = NzD(wsCfg.Cells(r, 34).Value, 0)
        kd3 = NzD(wsCfg.Cells(r, 35).Value, -1)
        m_KdProfiles(et) = Array(t0, kd0, t1, kd1, t2, kd2, t3, kd3)
NX:
    Next r
End Sub

Private Function EvalKdAbs(ByVal tSec As Double, ByVal prof As Variant) As Double
    Dim t(0 To 3) As Double, k(0 To 3) As Double
    Dim i As Long, n As Long, j As Long
    n = 0
    For i = 0 To 3
        If CDbl(prof(2 * i + 1)) >= 0 Then
            t(n) = CDbl(prof(2 * i))
            k(n) = CDbl(prof(2 * i + 1))
            n = n + 1
        End If
    Next i
    If n < 2 Then
        Err.Raise vbObjectError + 2403, , "Профиль Kд(t) должен содержать минимум 2 точки (AA:AI)."
    End If
    For j = 1 To n - 1
        If tSec <= t(j) Then
            EvalKdAbs = k(j - 1) + (k(j) - k(j - 1)) * SafeDiv((tSec - t(j - 1)), (t(j) - t(j - 1)), 0)
            Exit Function
        End If
    Next j
    EvalKdAbs = k(n - 1)
End Function

Private Function ClampKd(ByVal kdVal As Double) As Double
    ClampKd = kdVal
    If ClampKd < 0.1 Then ClampKd = 0.1
    If ClampKd > 1# Then ClampKd = 1#
End Function

' Сопоставить Тип_оборудования с ключом профиля AA:AI (допускаются короткие имена ПТУ/ПСУ/ГТУ).
Private Function ResolveKdKey(ByVal equipType As String) As String
    Dim key As String
    key = UCase$(Trim$(equipType))
    ResolveKdKey = key
    If m_KdProfiles Is Nothing Then Exit Function
    If m_KdProfiles.Exists(key) Then Exit Function
    If InStr(key, "ГПА") > 0 And m_KdProfiles.Exists("ГПА") Then ResolveKdKey = "ГПА": Exit Function
    If InStr(key, "ГТУ") > 0 And m_KdProfiles.Exists("ГТУ") Then ResolveKdKey = "ГТУ": Exit Function
    If (InStr(key, "ПГУ") > 0 And (InStr(key, "СБРОС") > 0 Or InStr(key, "ПАРАЛ") > 0)) And m_KdProfiles.Exists("ПГУ_СБРОСН") Then
        ResolveKdKey = "ПГУ_СБРОСН": Exit Function
    End If
    If InStr(key, "ПГУ") > 0 And m_KdProfiles.Exists("ПГУ_УТИЛ") Then ResolveKdKey = "ПГУ_УТИЛ": Exit Function
    If (InStr(key, "НЕБЛОК") > 0 Or InStr(key, "ОБЩ") > 0) And m_KdProfiles.Exists("ПТУ_НЕБЛОК") Then
        ResolveKdKey = "ПТУ_НЕБЛОК": Exit Function
    End If
    If (InStr(key, "УГОЛЬ") > 0 Or InStr(key, "ПЫЛЕУГ") > 0) And m_KdProfiles.Exists("ПТУ_БЛОК_УГОЛЬ") Then
        ResolveKdKey = "ПТУ_БЛОК_УГОЛЬ": Exit Function
    End If
    If (InStr(key, "ПТУ") > 0 Or InStr(key, "ПСУ") > 0 Or InStr(key, "БЛОК") > 0) And m_KdProfiles.Exists("ПТУ_БЛОК") Then
        ResolveKdKey = "ПТУ_БЛОК": Exit Function
    End If
End Function

Private Function DynamicKdByTime(ByVal equipType As String, ByVal tSec As Double) As Double
    Dim prof As Variant, key As String
    If m_KdProfiles Is Nothing Then Err.Raise vbObjectError + 2401, , "Профили Kд(t) не загружены."
    key = ResolveKdKey(equipType)
    If Not m_KdProfiles.Exists(key) Then
        Err.Raise vbObjectError + 2402, , "Для типа '" & equipType & "' не задан профиль Kд(t) в Config (AA:AI)."
    End If
    prof = m_KdProfiles(key)
    DynamicKdByTime = ClampKd(EvalKdAbs(tSec, prof))
End Function

Private Function KdProfileText(ByVal equipType As String) As String
    Dim prof As Variant, key As String, src As String
    src = "Config"
    If m_KdProfiles Is Nothing Then
        KdProfileText = "Config: профили Kд(t) не загружены"
        Exit Function
    End If
    key = ResolveKdKey(equipType)
    If Not m_KdProfiles.Exists(key) Then
        KdProfileText = "Config: профиль для типа '" & equipType & "' не задан"
        Exit Function
    End If
    prof = m_KdProfiles(key)
    KdProfileText = src & ": t0=" & Format(prof(0), "0") & "с Kд0=" & Format(prof(1), "0.00") & _
                    "; t1=" & Format(prof(2), "0") & "с Kд1=" & Format(prof(3), "0.00") & _
                    "; t2=" & Format(prof(4), "0") & "с Kд2=" & Format(prof(5), "0.00") & _
                    "; t3=" & Format(prof(6), "0") & "с Kд3=" & Format(prof(7), "0.00")
End Function

Private Sub LoadChartDisplayOptions(ByVal wsCfg As Worksheet)
    Dim r As Long
    Set m_ChartOpts = CreateObject("Scripting.Dictionary")
    For r = 3 To 10
        m_ChartOpts("REL_" & CStr(r)) = (NzD(wsCfg.Cells(r, 37).Value, 1) <> 0) ' AK
        m_ChartOpts("ABS_" & CStr(r)) = (NzD(wsCfg.Cells(r, 38).Value, 1) <> 0) ' AL
    Next r
    ' Row mapping: 3 Pfact, 4 Preq, 5 Tol+, 6 Tol-, 7 Limits, 8 Freq, 9 Markers
    m_ChartOpts("REL_PFACT") = m_ChartOpts("REL_3")
    m_ChartOpts("REL_PREQ") = m_ChartOpts("REL_4")
    m_ChartOpts("REL_TOL_P") = m_ChartOpts("REL_5")
    m_ChartOpts("REL_TOL_M") = m_ChartOpts("REL_6")
    m_ChartOpts("REL_LIMITS") = m_ChartOpts("REL_7")
    m_ChartOpts("REL_FREQ") = m_ChartOpts("REL_8")
    m_ChartOpts("REL_MARKERS") = m_ChartOpts("REL_9")
    m_ChartOpts("ABS_PFACT") = m_ChartOpts("ABS_3")
    m_ChartOpts("ABS_PREQ") = m_ChartOpts("ABS_4")
    m_ChartOpts("ABS_TOL_P") = m_ChartOpts("ABS_5")
    m_ChartOpts("ABS_TOL_M") = m_ChartOpts("ABS_6")
    m_ChartOpts("ABS_LIMITS") = m_ChartOpts("ABS_7")
    m_ChartOpts("ABS_FREQ") = m_ChartOpts("ABS_8")
    m_ChartOpts("ABS_MARKERS") = m_ChartOpts("ABS_9")
End Sub

Private Function ChartOpt(ByVal key As String, ByVal fallback As Boolean) As Boolean
    If m_ChartOpts Is Nothing Then
        ChartOpt = fallback
    ElseIf m_ChartOpts.Exists(key) Then
        ChartOpt = CBool(m_ChartOpts(key))
    Else
        ChartOpt = fallback
    End If
End Function

Private Sub EnsureChartOptionCheckboxes(ByVal ws As Worksheet)
    Dim r As Long
    Dim cb As CheckBox
    Dim nm As String, c As Long
    On Error Resume Next
    For Each cb In ws.CheckBoxes
        If Left$(cb.Name, 7) = "chkOpt_" Then cb.Delete
    Next cb
    On Error GoTo 0

    For r = 3 To 9
        For c = 37 To 38 ' AK, AL
            nm = "chkOpt_" & r & "_" & c
            Set cb = ws.CheckBoxes.Add(ws.Cells(r, c).Left + 2, ws.Cells(r, c).Top + 2, ws.Cells(r, c).Width - 4, ws.Cells(r, c).Height - 4)
            cb.Name = nm
            cb.Caption = ""
            cb.LinkedCell = ws.Cells(r, c).Address(True, True, xlA1, True)
            cb.Value = IIf(NzD(ws.Cells(r, c).Value, 1) <> 0, xlOn, xlOff)
            cb.Placement = xlMoveAndSize
        Next c
    Next r
End Sub

' ==========================================================
' Служебные
' ==========================================================

Private Sub CollectOldOutputSheets(ByRef names As Collection)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SH_RAW And ws.Name <> SH_CFG And ws.Name <> SH_SUM And ws.Name <> SH_LOG Then
            names.Add ws.Name
        End If
    Next ws
End Sub

Private Sub DeleteOutputSheets(ByVal names As Collection)
    Dim i As Long
    Application.DisplayAlerts = False
    For i = names.Count To 1 Step -1
        On Error Resume Next
        ThisWorkbook.Worksheets(CStr(names(i))).Delete
        On Error GoTo 0
    Next i
    Application.DisplayAlerts = True
End Sub

Private Function StationMatch(ByVal a As String, ByVal b As String) As Boolean
    Dim na As String, nb As String
    na = UCase$(Trim$(a))
    nb = UCase$(Trim$(b))
    If InStr(na, "ВОРКУТ") > 0 And InStr(nb, "ВОРКУТ") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "СОСНОГОР") > 0 And InStr(nb, "СОСНОГОР") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "СЛПК") > 0 And InStr(nb, "СЛПК") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "ПЕЧОР") > 0 And InStr(nb, "ПЕЧОР") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "УСИНС") > 0 And InStr(nb, "УСИНС") > 0 Then StationMatch = True: Exit Function
    If InStr(na, "ЯРЕГ") > 0 And InStr(nb, "ЯРЕГ") > 0 Then StationMatch = True: Exit Function
    StationMatch = (na = nb)
End Function

Private Function MinArray(ByRef arr() As Double) As Double
    Dim i As Long, v As Double
    v = 1E+99
    For i = LBound(arr) To UBound(arr)
        If arr(i) > 0 And arr(i) < v Then v = arr(i)
    Next i
    If v = 1E+99 Then v = 0.05
    MinArray = v
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then MaxD = a Else MaxD = b
End Function

Private Function FormatSecOrNA(ByVal secVal As Double) As String
    If secVal < 0 Then FormatSecOrNA = "н/д" Else FormatSecOrNA = Format(secVal, "0.0")
End Function

' Суффикс события: "" для первого, "_E2", "_E3"... для последующих
Private Function EventSuffix(ByVal evIdx As Long) As String
    If evIdx > 1 Then EventSuffix = "_E" & evIdx Else EventSuffix = ""
End Function

Private Function GeneratorSheetName(ByRef g As TGenCfg, ByVal evIdx As Long) As String
    Dim base As String, sfx As String
    sfx = EventSuffix(evIdx)
    base = g.Station & "_" & g.Generator
    If Len(base & sfx) > 31 Then base = Left$(base, 31 - Len(sfx))
    GeneratorSheetName = MakeSheetName(base & sfx)
End Function

Private Function GeneratorChartSheetName(ByRef g As TGenCfg, ByVal evIdx As Long) As String
    Dim base As String, full As String, sfx As String
    sfx = EventSuffix(evIdx)
    base = g.Station & "_" & g.Generator
    full = base & sfx & CHART_SUFFIX
    If Len(full) > 31 Then
        base = Left$(base, 31 - Len(CHART_SUFFIX) - Len(sfx))
        full = base & sfx & CHART_SUFFIX
    End If
    GeneratorChartSheetName = MakeSheetName(full)
End Function

Private Function MakeSheetName(ByVal s As String) As String
    Dim out As String
    out = s
    out = Replace(out, ":", "_")
    out = Replace(out, "\", "_")
    out = Replace(out, "/", "_")
    out = Replace(out, "?", "_")
    out = Replace(out, "*", "_")
    out = Replace(out, "[", "_")
    out = Replace(out, "]", "_")
    If Len(out) > 31 Then out = Left$(out, 31)
    If Len(out) = 0 Then out = "Sheet_OPRCH"
    MakeSheetName = out
End Function

Private Function GetRequiredSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set GetRequiredSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetRequiredSheet Is Nothing Then
        Err.Raise vbObjectError + 2999, , "Не найден лист '" & name & "'. Запустите SetupOPRCHTemplate."
    End If
End Function

Private Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
    End If
End Function

Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long, v As String
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        v = Trim$(CStr(ws.Cells(1, c).Value))
        If StrComp(v, headerName, vbTextCompare) = 0 Then FindHeaderCol = c: Exit Function
    Next c
    FindHeaderCol = 0
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If r < 1 Then r = 1
    LastUsedRow = r
End Function

Private Function NzD(ByVal v As Variant, Optional ByVal fallback As Double = 0#) As Double
    If IsError(v) Or IsEmpty(v) Or Trim$(CStr(v)) = "" Then
        NzD = fallback
    ElseIf IsNumeric(v) Then
        NzD = CDbl(v)
    Else
        NzD = fallback
    End If
End Function

Private Function SafeDiv(ByVal a As Double, ByVal b As Double, ByVal fallback As Double) As Double
    If Abs(b) < 0.0000001 Then SafeDiv = fallback Else SafeDiv = a / b
End Function

Private Function SgnNZ(ByVal v As Double) As Integer
    If v > 0.000001 Then
        SgnNZ = 1
    ElseIf v < -0.000001 Then
        SgnNZ = -1
    Else
        SgnNZ = 0
    End If
End Function

' ==========================================================
' Кнопки управления
' ==========================================================

Private Sub EnsureControlButtons(ByVal ws As Worksheet)
    AddOrReplaceButton ws, "btnRunOPRCH", ws.Range("E30"), 260, 32, _
        "Запустить мониторинг ОПРЧ", "AnalyzeOPRCH", RGB(40, 120, 220), RGB(255, 255, 255)
    AddOrReplaceButton ws, "btnClearOPRCH", ws.Range("I30"), 200, 32, _
        "Очистить результаты", "ClearOPRCHResults", RGB(150, 150, 150), RGB(255, 255, 255)
    AddOrReplaceButton ws, "btnPresets", ws.Range("M30"), 220, 32, _
        "Применить пресеты типов", "ApplyPresetsToConfig", RGB(80, 160, 100), RGB(255, 255, 255)
    AddOrReplaceButton ws, "btnExportPdf", ws.Range("Q30"), 200, 32, _
        "Экспорт отчёта в PDF", "ExportOPRCHReport", RGB(160, 90, 160), RGB(255, 255, 255)
End Sub

Private Sub AddOrReplaceButton(ByVal ws As Worksheet, ByVal nm As String, ByVal anchor As Range, _
                               ByVal w As Double, ByVal h As Double, _
                               ByVal caption As String, ByVal onAction As String, _
                               ByVal fillColor As Long, ByVal textColor As Long)
    Dim shp As Shape
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, w, h)
    shp.Name = nm
    shp.TextFrame2.TextRange.Text = caption
    shp.OnAction = onAction
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(70, 70, 70)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = textColor
    shp.TextFrame2.TextRange.Font.Size = 11
End Sub
