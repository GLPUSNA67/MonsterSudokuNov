Public Class frmSudokuMonster
    Public Property CurrentRow As Int16
    Public Property CurrentColumn As Int16
    Public Property CurrentSquare As Int16
    Public Property CurrentNumber As String
    Public Property Hints As String
    Public Property StartingString As String
    Public Property bRemoveANumberFromACell As Boolean
    Public Property bAutoComplete As Boolean
    Public Property D As Int16
    Public Property R As Int16
    Public Property NullValue As String
    Dim arrDone(255) As Boolean
    Dim arrDone0(255) As Boolean
    Dim arrDone1(255) As Boolean
    Dim arrDone2(255) As Boolean
    Dim bShowRemainders As Boolean
    Dim bInitialShowRemainders As Boolean
    Dim arrValues(255) As String
    Dim arrValues0(255) As String
    Dim arrValues1(255) As String
    Dim arrValues2(255) As String

    Dim arrReferences(255, 3) As Int16

    Dim arrStringsRemaining(2, 16) As String
    Dim bShow As Boolean

    Private Sub frmMonsterSudoku_Load(sender As Object, e As EventArgs)

        'TextBox1.Text = TextBox1.Text & "Entering frmMonsterSudoku_Load " & vbCrLf

        StartingString = "0123456789ABCDEF"
        Dim i As Int16

        bShow = True
        bRemoveANumberFromACell = False

        'Initialize arrays arrDone to False and arrValues to the starting string
        For i = 0 To 255
            arrDone(i) = False
            arrValues(i) = StartingString
            'arrReferences(i, 3) = StartingString   'an alternate way to store all information (except done) in one array
        Next i

        arrReferences(0, 0) = 1
        'Me.TextBox1.Text = Me.TextBox1.Text & "arrReferences(0, 0) is " & arrReferences(0, 0) & vbCrLf
        arrReferences(0, 1) = 1
        arrReferences(0, 2) = 1
        arrReferences(1, 0) = 1
        arrReferences(1, 1) = 2
        arrReferences(1, 2) = 1
        arrReferences(2, 0) = 1
        arrReferences(2, 1) = 3
        arrReferences(2, 2) = 1
        arrReferences(3, 0) = 1
        arrReferences(3, 1) = 4
        arrReferences(3, 2) = 1
        arrReferences(4, 0) = 1
        arrReferences(4, 1) = 5
        arrReferences(4, 2) = 2
        arrReferences(5, 0) = 1
        arrReferences(5, 1) = 6
        arrReferences(5, 2) = 2
        arrReferences(6, 0) = 1
        arrReferences(6, 1) = 7
        arrReferences(6, 2) = 2
        arrReferences(7, 0) = 1
        arrReferences(7, 1) = 7
        arrReferences(7, 2) = 2
        arrReferences(8, 0) = 1
        arrReferences(8, 1) = 9
        arrReferences(8, 2) = 3
        arrReferences(9, 0) = 1
        arrReferences(9, 1) = 10
        arrReferences(9, 2) = 3
        arrReferences(10, 0) = 1
        arrReferences(10, 1) = 11
        arrReferences(10, 2) = 3
        arrReferences(11, 0) = 1
        arrReferences(11, 1) = 12
        arrReferences(11, 2) = 3
        arrReferences(12, 0) = 1
        arrReferences(12, 1) = 13
        arrReferences(12, 2) = 4
        arrReferences(13, 0) = 1
        arrReferences(13, 1) = 14
        arrReferences(13, 2) = 4
        arrReferences(14, 0) = 1
        arrReferences(14, 1) = 15
        arrReferences(14, 2) = 4
        arrReferences(15, 0) = 1
        arrReferences(15, 1) = 16
        arrReferences(15, 2) = 4
        arrReferences(16, 0) = "2"
        arrReferences(16, 1) = "1"
        arrReferences(16, 2) = "1"
        arrReferences(17, 0) = "2"
        arrReferences(17, 1) = "2"
        arrReferences(17, 2) = "1"
        arrReferences(18, 0) = "2"
        arrReferences(18, 1) = "3"
        arrReferences(18, 2) = "1"
        arrReferences(19, 0) = "2"
        arrReferences(19, 1) = "4"
        arrReferences(19, 2) = "1"
        arrReferences(20, 0) = "2"
        arrReferences(20, 1) = "5"
        arrReferences(20, 2) = "2"
        arrReferences(21, 0) = "2"
        arrReferences(21, 1) = "6"
        arrReferences(21, 2) = "2"
        arrReferences(22, 0) = "2"
        arrReferences(22, 1) = "7"
        arrReferences(22, 2) = "2"
        arrReferences(23, 0) = "2"
        arrReferences(23, 1) = "8"
        arrReferences(23, 2) = "2"
        arrReferences(24, 0) = "2"
        arrReferences(24, 1) = "9"
        arrReferences(24, 2) = "3"
        arrReferences(25, 0) = "2"
        arrReferences(25, 1) = "10"
        arrReferences(25, 2) = "3"
        arrReferences(26, 0) = "2"
        arrReferences(26, 1) = "11"
        arrReferences(26, 2) = "3"
        arrReferences(27, 0) = "2"
        arrReferences(27, 1) = "12"
        arrReferences(27, 2) = "3"
        arrReferences(28, 0) = "2"
        arrReferences(28, 1) = "13"
        arrReferences(28, 2) = "4"
        arrReferences(29, 0) = "2"
        arrReferences(29, 1) = "14"
        arrReferences(29, 2) = "4"
        arrReferences(30, 0) = "2"
        arrReferences(30, 1) = "15"
        arrReferences(30, 2) = "4"
        arrReferences(31, 0) = "2"
        arrReferences(31, 1) = "16"
        arrReferences(31, 2) = "4"
        arrReferences(32, 0) = "3"
        arrReferences(32, 1) = "1"
        arrReferences(32, 2) = "1"
        arrReferences(33, 0) = "3"
        arrReferences(33, 1) = "2"
        arrReferences(33, 2) = "1"
        arrReferences(34, 0) = "3"
        arrReferences(34, 1) = "3"
        arrReferences(34, 2) = "1"
        arrReferences(35, 0) = "3"
        arrReferences(35, 1) = "4"
        arrReferences(35, 2) = "1"
        arrReferences(36, 0) = "3"
        arrReferences(36, 1) = "5"
        arrReferences(36, 2) = "2"
        arrReferences(37, 0) = "3"
        arrReferences(37, 1) = "6"
        arrReferences(37, 2) = "2"
        arrReferences(38, 0) = "3"
        arrReferences(38, 1) = "7"
        arrReferences(38, 2) = "2"
        arrReferences(39, 0) = "3"
        arrReferences(39, 1) = "8"
        arrReferences(39, 2) = "2"
        arrReferences(40, 0) = "3"
        arrReferences(40, 1) = "9"
        arrReferences(40, 2) = "3"
        arrReferences(41, 0) = "3"
        arrReferences(41, 1) = "10"
        arrReferences(41, 2) = "3"
        arrReferences(42, 0) = "3"
        arrReferences(42, 1) = "11"
        arrReferences(42, 2) = "3"
        arrReferences(43, 0) = "3"
        arrReferences(43, 1) = "12"
        arrReferences(43, 2) = "3"
        arrReferences(44, 0) = "3"
        arrReferences(44, 1) = "13"
        arrReferences(44, 2) = "4"
        arrReferences(45, 0) = "3"
        arrReferences(45, 1) = "14"
        arrReferences(45, 2) = "4"
        arrReferences(46, 0) = "3"
        arrReferences(46, 1) = "15"
        arrReferences(46, 2) = "4"
        arrReferences(47, 0) = "3"
        arrReferences(47, 1) = "16"
        arrReferences(47, 2) = "4"
        arrReferences(48, 0) = "4"
        arrReferences(48, 1) = "1"
        arrReferences(48, 2) = "1"
        arrReferences(49, 0) = "4"
        arrReferences(49, 1) = "2"
        arrReferences(49, 2) = "1"
        arrReferences(50, 0) = "4"
        arrReferences(50, 1) = "3"
        arrReferences(50, 2) = "1"
        arrReferences(51, 0) = "4"
        arrReferences(51, 1) = "4"
        arrReferences(51, 2) = "1"
        arrReferences(52, 0) = "4"
        arrReferences(52, 1) = "5"
        arrReferences(52, 2) = "2"
        arrReferences(53, 0) = "4"
        arrReferences(53, 1) = "6"
        arrReferences(53, 2) = "2"
        arrReferences(54, 0) = "4"
        arrReferences(54, 1) = "7"
        arrReferences(54, 2) = "2"
        arrReferences(55, 0) = "4"
        arrReferences(55, 1) = "8"
        arrReferences(55, 2) = "2"
        arrReferences(56, 0) = "4"
        arrReferences(56, 1) = "9"
        arrReferences(56, 2) = "3"
        arrReferences(57, 0) = "4"
        arrReferences(57, 1) = "10"
        arrReferences(57, 2) = "3"
        arrReferences(58, 0) = "4"
        arrReferences(58, 1) = "11"
        arrReferences(58, 2) = "3"
        arrReferences(59, 0) = "4"
        arrReferences(59, 1) = "12"
        arrReferences(59, 2) = "3"
        arrReferences(60, 0) = "4"
        arrReferences(60, 1) = "13"
        arrReferences(60, 2) = "4"
        arrReferences(61, 0) = "4"
        arrReferences(61, 1) = "14"
        arrReferences(61, 2) = "4"
        arrReferences(62, 0) = "4"
        arrReferences(62, 1) = "15"
        arrReferences(62, 2) = "4"
        arrReferences(63, 0) = "4"
        arrReferences(63, 1) = "16"
        arrReferences(63, 2) = "4"
        arrReferences(64, 0) = "5"
        arrReferences(64, 1) = "1"
        arrReferences(64, 2) = "5"
        arrReferences(65, 0) = "5"
        arrReferences(65, 1) = "2"
        arrReferences(65, 2) = "5"
        arrReferences(66, 0) = "5"
        arrReferences(66, 1) = "3"
        arrReferences(66, 2) = "5"
        arrReferences(67, 0) = "5"
        arrReferences(67, 1) = "4"
        arrReferences(67, 2) = "5"
        arrReferences(68, 0) = "5"
        arrReferences(68, 1) = "5"
        arrReferences(68, 2) = "6"
        arrReferences(69, 0) = "5"
        arrReferences(69, 1) = "6"
        arrReferences(69, 2) = "6"
        arrReferences(70, 0) = "5"
        arrReferences(70, 1) = "7"
        arrReferences(70, 2) = "6"
        arrReferences(71, 0) = "5"
        arrReferences(71, 1) = "8"
        arrReferences(71, 2) = "6"
        arrReferences(72, 0) = "5"
        arrReferences(72, 1) = "9"
        arrReferences(72, 2) = "7"
        arrReferences(73, 0) = "5"
        arrReferences(73, 1) = "10"
        arrReferences(73, 2) = "7"
        arrReferences(74, 0) = "5"
        arrReferences(74, 1) = "11"
        arrReferences(74, 2) = "7"
        arrReferences(75, 0) = "5"
        arrReferences(75, 1) = "12"
        arrReferences(75, 2) = "7"
        arrReferences(76, 0) = "5"
        arrReferences(76, 1) = "13"
        arrReferences(76, 2) = "8"
        arrReferences(77, 0) = "5"
        arrReferences(77, 1) = "14"
        arrReferences(77, 2) = "8"
        arrReferences(78, 0) = "5"
        arrReferences(78, 1) = "15"
        arrReferences(78, 2) = "8"
        arrReferences(79, 0) = "5"
        arrReferences(79, 1) = "16"
        arrReferences(79, 2) = "8"
        arrReferences(80, 0) = "6"
        arrReferences(80, 1) = "1"
        arrReferences(80, 2) = "5"
        arrReferences(81, 0) = "6"
        arrReferences(81, 1) = "2"
        arrReferences(81, 2) = "5"
        arrReferences(82, 0) = "6"
        arrReferences(82, 1) = "3"
        arrReferences(82, 2) = "5"
        arrReferences(83, 0) = "6"
        arrReferences(83, 1) = "4"
        arrReferences(83, 2) = "5"
        arrReferences(84, 0) = "6"
        arrReferences(84, 1) = "5"
        arrReferences(84, 2) = "6"
        arrReferences(85, 0) = "6"
        arrReferences(85, 1) = "6"
        arrReferences(85, 2) = "6"
        arrReferences(86, 0) = "6"
        arrReferences(86, 1) = "7"
        arrReferences(86, 2) = "6"
        arrReferences(87, 0) = "6"
        arrReferences(87, 1) = "8"
        arrReferences(87, 2) = "6"
        arrReferences(88, 0) = "6"
        arrReferences(88, 1) = "9"
        arrReferences(88, 2) = "7"
        arrReferences(89, 0) = "6"
        arrReferences(89, 1) = "10"
        arrReferences(89, 2) = "7"
        arrReferences(90, 0) = "6"
        arrReferences(90, 1) = "11"
        arrReferences(90, 2) = "7"
        arrReferences(91, 0) = "6"
        arrReferences(91, 1) = "12"
        arrReferences(91, 2) = "7"
        arrReferences(92, 0) = "6"
        arrReferences(92, 1) = "13"
        arrReferences(92, 2) = "8"
        arrReferences(93, 0) = "6"
        arrReferences(93, 1) = "14"
        arrReferences(93, 2) = "8"
        arrReferences(94, 0) = "6"
        arrReferences(94, 1) = "15"
        arrReferences(94, 2) = "8"
        arrReferences(95, 0) = "6"
        arrReferences(95, 1) = "16"
        arrReferences(95, 2) = "8"
        arrReferences(96, 0) = "7"
        arrReferences(96, 1) = "1"
        arrReferences(96, 2) = "5"
        arrReferences(97, 0) = "7"
        arrReferences(97, 1) = "2"
        arrReferences(97, 2) = "5"
        arrReferences(98, 0) = "7"
        arrReferences(98, 1) = "3"
        arrReferences(98, 2) = "5"
        arrReferences(99, 0) = "7"
        arrReferences(99, 1) = "4"
        arrReferences(99, 2) = "5"
        arrReferences(100, 0) = "7"
        arrReferences(100, 1) = "5"
        arrReferences(100, 2) = "6"
        arrReferences(101, 0) = "7"
        arrReferences(101, 1) = "6"
        arrReferences(101, 2) = "6"
        arrReferences(102, 0) = "7"
        arrReferences(102, 1) = "7"
        arrReferences(102, 2) = "6"
        arrReferences(103, 0) = "7"
        arrReferences(103, 1) = "8"
        arrReferences(103, 2) = "6"
        arrReferences(104, 0) = "7"
        arrReferences(104, 1) = "9"
        arrReferences(104, 2) = "7"
        arrReferences(105, 0) = "7"
        arrReferences(105, 1) = "10"
        arrReferences(105, 2) = "7"
        arrReferences(106, 0) = "7"
        arrReferences(106, 1) = "11"
        arrReferences(106, 2) = "7"
        arrReferences(107, 0) = "7"
        arrReferences(107, 1) = "12"
        arrReferences(107, 2) = "7"
        arrReferences(108, 0) = "7"
        arrReferences(108, 1) = "13"
        arrReferences(108, 2) = "8"
        arrReferences(109, 0) = "7"
        arrReferences(109, 1) = "14"
        arrReferences(109, 2) = "8"
        arrReferences(110, 0) = "7"
        arrReferences(110, 1) = "15"
        arrReferences(110, 2) = "8"
        arrReferences(111, 0) = "7"
        arrReferences(111, 1) = "16"
        arrReferences(111, 2) = "8"
        arrReferences(112, 0) = "8"
        arrReferences(112, 1) = "1"
        arrReferences(112, 2) = "5"
        arrReferences(113, 0) = "8"
        arrReferences(113, 1) = "2"
        arrReferences(113, 2) = "5"
        arrReferences(114, 0) = "8"
        arrReferences(114, 1) = "3"
        arrReferences(114, 2) = "5"
        arrReferences(115, 0) = "8"
        arrReferences(115, 1) = "4"
        arrReferences(115, 2) = "5"
        arrReferences(116, 0) = "8"
        arrReferences(116, 1) = "5"
        arrReferences(116, 2) = "6"
        arrReferences(117, 0) = "8"
        arrReferences(117, 1) = "6"
        arrReferences(117, 2) = "6"
        arrReferences(118, 0) = "8"
        arrReferences(118, 1) = "7"
        arrReferences(118, 2) = "6"
        arrReferences(119, 0) = "8"
        arrReferences(119, 1) = "8"
        arrReferences(119, 2) = "6"
        arrReferences(120, 0) = "8"
        arrReferences(120, 1) = "9"
        arrReferences(120, 2) = "7"
        arrReferences(121, 0) = "8"
        arrReferences(121, 1) = "10"
        arrReferences(121, 2) = "7"
        arrReferences(122, 0) = "8"
        arrReferences(122, 1) = "11"
        arrReferences(122, 2) = "7"
        arrReferences(123, 0) = "8"
        arrReferences(123, 1) = "12"
        arrReferences(123, 2) = "7"
        arrReferences(124, 0) = "8"
        arrReferences(124, 1) = "13"
        arrReferences(124, 2) = "8"
        arrReferences(125, 0) = "8"
        arrReferences(125, 1) = "14"
        arrReferences(125, 2) = "8"
        arrReferences(126, 0) = "8"
        arrReferences(126, 1) = "15"
        arrReferences(126, 2) = "8"
        arrReferences(127, 0) = "8"
        arrReferences(127, 1) = "16"
        arrReferences(127, 2) = "8"
        arrReferences(128, 0) = "9"
        arrReferences(128, 1) = "1"
        arrReferences(128, 2) = "9"
        arrReferences(129, 0) = "9"
        arrReferences(129, 1) = "2"
        arrReferences(129, 2) = "9"
        arrReferences(130, 0) = "9"
        arrReferences(130, 1) = "3"
        arrReferences(130, 2) = "9"
        arrReferences(131, 0) = "9"
        arrReferences(131, 1) = "4"
        arrReferences(131, 2) = "9"
        arrReferences(132, 0) = "9"
        arrReferences(132, 1) = "5"
        arrReferences(132, 2) = "10"
        arrReferences(133, 0) = "9"
        arrReferences(133, 1) = "6"
        arrReferences(133, 2) = "10"
        arrReferences(134, 0) = "9"
        arrReferences(134, 1) = "7"
        arrReferences(134, 2) = "10"
        arrReferences(135, 0) = "9"
        arrReferences(135, 1) = "8"
        arrReferences(135, 2) = "10"
        arrReferences(136, 0) = "9"
        arrReferences(136, 1) = "9"
        arrReferences(136, 2) = "11"
        arrReferences(137, 0) = "9"
        arrReferences(137, 1) = "10"
        arrReferences(137, 2) = "11"
        arrReferences(138, 0) = "9"
        arrReferences(138, 1) = "11"
        arrReferences(138, 2) = "11"
        arrReferences(139, 0) = "9"
        arrReferences(139, 1) = "12"
        arrReferences(139, 2) = "11"
        arrReferences(140, 0) = "9"
        arrReferences(140, 1) = "13"
        arrReferences(140, 2) = "12"
        arrReferences(141, 0) = "9"
        arrReferences(141, 1) = "14"
        arrReferences(141, 2) = "12"
        arrReferences(142, 0) = "9"
        arrReferences(142, 1) = "15"
        arrReferences(142, 2) = "12"
        arrReferences(143, 0) = "9"
        arrReferences(143, 1) = "16"
        arrReferences(143, 2) = "12"
        arrReferences(144, 0) = "10"
        arrReferences(144, 1) = "1"
        arrReferences(144, 2) = "9"
        arrReferences(145, 0) = "10"
        arrReferences(145, 1) = "2"
        arrReferences(145, 2) = "9"
        arrReferences(146, 0) = "10"
        arrReferences(146, 1) = "3"
        arrReferences(146, 2) = "9"
        arrReferences(147, 0) = "10"
        arrReferences(147, 1) = "4"
        arrReferences(147, 2) = "9"
        arrReferences(148, 0) = "10"
        arrReferences(148, 1) = "5"
        arrReferences(148, 2) = "10"
        arrReferences(149, 0) = "10"
        arrReferences(149, 1) = "6"
        arrReferences(149, 2) = "10"
        arrReferences(150, 0) = "10"
        arrReferences(150, 1) = "7"
        arrReferences(150, 2) = "10"
        arrReferences(151, 0) = "10"
        arrReferences(151, 1) = "8"
        arrReferences(151, 2) = "10"
        arrReferences(152, 0) = "10"
        arrReferences(152, 1) = "9"
        arrReferences(152, 2) = "11"
        arrReferences(153, 0) = "10"
        arrReferences(153, 1) = "10"
        arrReferences(153, 2) = "11"
        arrReferences(154, 0) = "10"
        arrReferences(154, 1) = "11"
        arrReferences(154, 2) = "11"
        arrReferences(155, 0) = "10"
        arrReferences(155, 1) = "12"
        arrReferences(155, 2) = "11"
        arrReferences(156, 0) = "10"
        arrReferences(156, 1) = "13"
        arrReferences(156, 2) = "12"
        arrReferences(157, 0) = "10"
        arrReferences(157, 1) = "14"
        arrReferences(157, 2) = "12"
        arrReferences(158, 0) = "10"
        arrReferences(158, 1) = "15"
        arrReferences(158, 2) = "12"
        arrReferences(159, 0) = "10"
        arrReferences(159, 1) = "16"
        arrReferences(159, 2) = "12"
        arrReferences(160, 0) = "11"
        arrReferences(160, 1) = "1"
        arrReferences(160, 2) = "9"
        arrReferences(161, 0) = "11"
        arrReferences(161, 1) = "2"
        arrReferences(161, 2) = "9"
        arrReferences(162, 0) = "11"
        arrReferences(162, 1) = "3"
        arrReferences(162, 2) = "9"
        arrReferences(163, 0) = "11"
        arrReferences(163, 1) = "4"
        arrReferences(163, 2) = "9"
        arrReferences(164, 0) = "11"
        arrReferences(164, 1) = "5"
        arrReferences(164, 2) = "10"
        arrReferences(165, 0) = "11"
        arrReferences(165, 1) = "6"
        arrReferences(165, 2) = "10"
        arrReferences(166, 0) = "11"
        arrReferences(166, 1) = "7"
        arrReferences(166, 2) = "10"
        arrReferences(167, 0) = "11"
        arrReferences(167, 1) = "8"
        arrReferences(167, 2) = "10"
        arrReferences(168, 0) = "11"
        arrReferences(168, 1) = "9"
        arrReferences(168, 2) = "11"
        arrReferences(169, 0) = "11"
        arrReferences(169, 1) = "10"
        arrReferences(169, 2) = "11"
        arrReferences(170, 0) = "11"
        arrReferences(170, 1) = "11"
        arrReferences(170, 2) = "11"
        arrReferences(171, 0) = "11"
        arrReferences(171, 1) = "12"
        arrReferences(171, 2) = "11"
        arrReferences(172, 0) = "11"
        arrReferences(172, 1) = "13"
        arrReferences(172, 2) = "12"
        arrReferences(173, 0) = "11"
        arrReferences(173, 1) = "14"
        arrReferences(173, 2) = "12"
        arrReferences(174, 0) = "11"
        arrReferences(174, 1) = "15"
        arrReferences(174, 2) = "12"
        arrReferences(175, 0) = "11"
        arrReferences(175, 1) = "16"
        arrReferences(175, 2) = "12"
        arrReferences(176, 0) = "12"
        arrReferences(176, 1) = "1"
        arrReferences(176, 2) = "9"
        arrReferences(177, 0) = "12"
        arrReferences(177, 1) = "2"
        arrReferences(177, 2) = "9"
        arrReferences(178, 0) = "12"
        arrReferences(178, 1) = "3"
        arrReferences(178, 2) = "9"
        arrReferences(179, 0) = "12"
        arrReferences(179, 1) = "4"
        arrReferences(179, 2) = "9"
        arrReferences(180, 0) = "12"
        arrReferences(180, 1) = "5"
        arrReferences(180, 2) = "10"
        arrReferences(181, 0) = "12"
        arrReferences(181, 1) = "6"
        arrReferences(181, 2) = "10"
        arrReferences(182, 0) = "12"
        arrReferences(182, 1) = "7"
        arrReferences(182, 2) = "10"
        arrReferences(183, 0) = "12"
        arrReferences(183, 1) = "8"
        arrReferences(183, 2) = "10"
        arrReferences(184, 0) = "12"
        arrReferences(184, 1) = "9"
        arrReferences(184, 2) = "11"
        arrReferences(185, 0) = "12"
        arrReferences(185, 1) = "10"
        arrReferences(185, 2) = "11"
        arrReferences(186, 0) = "12"
        arrReferences(186, 1) = "11"
        arrReferences(186, 2) = "11"
        arrReferences(187, 0) = "12"
        arrReferences(187, 1) = "12"
        arrReferences(187, 2) = "11"
        arrReferences(188, 0) = "12"
        arrReferences(188, 1) = "13"
        arrReferences(188, 2) = "12"
        arrReferences(189, 0) = "12"
        arrReferences(189, 1) = "14"
        arrReferences(189, 2) = "12"
        arrReferences(190, 0) = "12"
        arrReferences(190, 1) = "15"
        arrReferences(190, 2) = "12"
        arrReferences(191, 0) = "12"
        arrReferences(191, 1) = "16"
        arrReferences(191, 2) = "12"
        arrReferences(192, 0) = "13"
        arrReferences(192, 1) = "1"
        arrReferences(192, 2) = "13"
        arrReferences(193, 0) = "13"
        arrReferences(193, 1) = "2"
        arrReferences(193, 2) = "13"
        arrReferences(194, 0) = "13"
        arrReferences(194, 1) = "3"
        arrReferences(194, 2) = "13"
        arrReferences(195, 0) = "13"
        arrReferences(195, 1) = "4"
        arrReferences(195, 2) = "13"
        arrReferences(196, 0) = "13"
        arrReferences(196, 1) = "5"
        arrReferences(196, 2) = "14"
        arrReferences(197, 0) = "13"
        arrReferences(197, 1) = "6"
        arrReferences(197, 2) = "14"
        arrReferences(198, 0) = "13"
        arrReferences(198, 1) = "7"
        arrReferences(198, 2) = "14"
        arrReferences(199, 0) = "13"
        arrReferences(199, 1) = "8"
        arrReferences(199, 2) = "14"
        arrReferences(200, 0) = "13"
        arrReferences(200, 1) = "9"
        arrReferences(200, 2) = "15"
        arrReferences(201, 0) = "13"
        arrReferences(201, 1) = "10"
        arrReferences(201, 2) = "15"
        arrReferences(202, 0) = "13"
        arrReferences(202, 1) = "11"
        arrReferences(202, 2) = "15"
        arrReferences(203, 0) = "13"
        arrReferences(203, 1) = "12"
        arrReferences(203, 2) = "15"
        arrReferences(204, 0) = "13"
        arrReferences(204, 1) = "13"
        arrReferences(204, 2) = "16"
        arrReferences(205, 0) = "13"
        arrReferences(205, 1) = "14"
        arrReferences(205, 2) = "16"
        arrReferences(206, 0) = "13"
        arrReferences(206, 1) = "15"
        arrReferences(206, 2) = "16"
        arrReferences(207, 0) = "13"
        arrReferences(207, 1) = "16"
        arrReferences(207, 2) = "16"
        arrReferences(208, 0) = "14"
        arrReferences(208, 1) = "1"
        arrReferences(208, 2) = "13"
        arrReferences(209, 0) = "14"
        arrReferences(209, 1) = "2"
        arrReferences(209, 2) = "13"
        arrReferences(210, 0) = "14"
        arrReferences(210, 1) = "3"
        arrReferences(210, 2) = "13"
        arrReferences(211, 0) = "14"
        arrReferences(211, 1) = "4"
        arrReferences(211, 2) = "13"
        arrReferences(212, 0) = "14"
        arrReferences(212, 1) = "5"
        arrReferences(212, 2) = "14"
        arrReferences(213, 0) = "14"
        arrReferences(213, 1) = "6"
        arrReferences(213, 2) = "14"
        arrReferences(214, 0) = "14"
        arrReferences(214, 1) = "7"
        arrReferences(214, 2) = "14"
        arrReferences(215, 0) = "14"
        arrReferences(215, 1) = "8"
        arrReferences(215, 2) = "14"
        arrReferences(216, 0) = "14"
        arrReferences(216, 1) = "9"
        arrReferences(216, 2) = "15"
        arrReferences(217, 0) = "14"
        arrReferences(217, 1) = "10"
        arrReferences(217, 2) = "15"
        arrReferences(218, 0) = "14"
        arrReferences(218, 1) = "11"
        arrReferences(218, 2) = "15"
        arrReferences(219, 0) = "14"
        arrReferences(219, 1) = "12"
        arrReferences(219, 2) = "15"
        arrReferences(220, 0) = "14"
        arrReferences(220, 1) = "13"
        arrReferences(220, 2) = "16"
        arrReferences(221, 0) = "14"
        arrReferences(221, 1) = "14"
        arrReferences(221, 2) = "16"
        arrReferences(222, 0) = "14"
        arrReferences(222, 1) = "15"
        arrReferences(222, 2) = "16"
        arrReferences(223, 0) = "14"
        arrReferences(223, 1) = "16"
        arrReferences(223, 2) = "16"
        arrReferences(224, 0) = "15"
        arrReferences(224, 1) = "1"
        arrReferences(224, 2) = "13"
        arrReferences(225, 0) = "15"
        arrReferences(225, 1) = "2"
        arrReferences(225, 2) = "13"
        arrReferences(226, 0) = "15"
        arrReferences(226, 1) = "3"
        arrReferences(226, 2) = "13"
        arrReferences(227, 0) = "15"
        arrReferences(227, 1) = "4"
        arrReferences(227, 2) = "13"
        arrReferences(228, 0) = "15"
        arrReferences(228, 1) = "5"
        arrReferences(228, 2) = "14"
        arrReferences(229, 0) = "15"
        arrReferences(229, 1) = "6"
        arrReferences(229, 2) = "14"
        arrReferences(230, 0) = "15"
        arrReferences(230, 1) = "7"
        arrReferences(230, 2) = "14"
        arrReferences(231, 0) = "15"
        arrReferences(231, 1) = "8"
        arrReferences(231, 2) = "14"
        arrReferences(232, 0) = "15"
        arrReferences(232, 1) = "9"
        arrReferences(232, 2) = "15"
        arrReferences(233, 0) = "15"
        arrReferences(233, 1) = "10"
        arrReferences(233, 2) = "15"
        arrReferences(234, 0) = "15"
        arrReferences(234, 1) = "11"
        arrReferences(234, 2) = "15"
        arrReferences(235, 0) = "15"
        arrReferences(235, 1) = "12"
        arrReferences(235, 2) = "15"
        arrReferences(236, 0) = "15"
        arrReferences(236, 1) = "13"
        arrReferences(236, 2) = "16"
        arrReferences(237, 0) = "15"
        arrReferences(237, 1) = "14"
        arrReferences(237, 2) = "16"
        arrReferences(238, 0) = "15"
        arrReferences(238, 1) = "15"
        arrReferences(238, 2) = "16"
        arrReferences(239, 0) = "15"
        arrReferences(239, 1) = "16"
        arrReferences(239, 2) = "16"
        arrReferences(240, 0) = "16"
        arrReferences(240, 1) = "1"
        arrReferences(240, 2) = "13"
        arrReferences(241, 0) = "16"
        arrReferences(241, 1) = "2"
        arrReferences(241, 2) = "13"
        arrReferences(242, 0) = "16"
        arrReferences(242, 1) = "3"
        arrReferences(242, 2) = "13"
        arrReferences(243, 0) = "16"
        arrReferences(243, 1) = "4"
        arrReferences(243, 2) = "13"
        arrReferences(244, 0) = "16"
        arrReferences(244, 1) = "5"
        arrReferences(244, 2) = "14"
        arrReferences(245, 0) = "16"
        arrReferences(245, 1) = "6"
        arrReferences(245, 2) = "14"
        arrReferences(246, 0) = "16"
        arrReferences(246, 1) = "7"
        arrReferences(246, 2) = "14"
        arrReferences(247, 0) = "16"
        arrReferences(247, 1) = "8"
        arrReferences(247, 2) = "14"
        arrReferences(248, 0) = "16"
        arrReferences(248, 1) = "9"
        arrReferences(248, 2) = "15"
        arrReferences(249, 0) = "16"
        arrReferences(249, 1) = "10"
        arrReferences(249, 2) = "15"
        arrReferences(250, 0) = "16"
        arrReferences(250, 1) = "11"
        arrReferences(250, 2) = "15"
        arrReferences(251, 0) = "16"
        arrReferences(251, 1) = "12"
        arrReferences(251, 2) = "15"
        arrReferences(252, 0) = "16"
        arrReferences(252, 1) = "13"
        arrReferences(252, 2) = "16"
        arrReferences(253, 0) = "16"
        arrReferences(253, 1) = "14"
        arrReferences(253, 2) = "16"
        arrReferences(254, 0) = "16"
        arrReferences(254, 1) = "15"
        arrReferences(254, 2) = "16"
        arrReferences(255, 0) = "16"
        arrReferences(255, 1) = "16"
        arrReferences(255, 2) = "16"

        RefreshPuzzle()

    End Sub
    Private Sub Button0_Click(sender As Object, e As EventArgs) Handles Button0.Click
        CurrentNumber = "0"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CurrentNumber = "1"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        CurrentNumber = "2"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        CurrentNumber = "3"
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        CurrentNumber = "4"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        CurrentNumber = "5"
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        CurrentNumber = "6"
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        CurrentNumber = "7"
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        CurrentNumber = "8"
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        CurrentNumber = "9"
    End Sub

    Private Sub ButtonA_Click(sender As Object, e As EventArgs) Handles ButtonA.Click
        CurrentNumber = "A"
    End Sub

    Private Sub ButtonB_Click(sender As Object, e As EventArgs) Handles ButtonB.Click
        CurrentNumber = "B"
    End Sub

    Private Sub ButtonC_Click(sender As Object, e As EventArgs) Handles ButtonC.Click
        CurrentNumber = "C"
    End Sub

    Private Sub ButtonD_Click(sender As Object, e As EventArgs) Handles ButtonD.Click
        CurrentNumber = "D"
    End Sub

    Private Sub ButtonE_Click(sender As Object, e As EventArgs) Handles ButtonE.Click
        CurrentNumber = "E"
    End Sub

    Private Sub ButtonF_Click(sender As Object, e As EventArgs) Handles ButtonF.Click
        CurrentNumber = "F"
    End Sub

    Private Sub R1C1_Click(sender As Object, e As EventArgs) Handles R1C1.Click

        CurrentRow = "1"
        CurrentColumn = "1"
        CurrentSquare = "1"
        If bRemoveANumberFromACell = False Then
            R1C1.Text = CurrentNumber
            arrDone(0) = True
            arrValues(0) = CurrentNumber
        End If

        'R1C1.Font.Bold = True
        ''TextBox1.Text = TextBox1.Text & "CurrentRow = " & CurrentRow & vbCrLf
        ''TextBox1.Text = TextBox1.Text & "CurrentColumn = " & CurrentColumn & vbCrLf
        ''TextBox1.Text = TextBox1.Text & "CurrentSquare = " & CurrentSquare & vbCrLf
    End Sub

    Private Sub R1C2_Click(sender As Object, e As EventArgs) Handles R1C2.Click

        CurrentRow = "1"
        CurrentColumn = "2"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R1C2.Text = CurrentNumber
            arrDone(1) = True
            arrValues(1) = CurrentNumber
        End If

    End Sub

    Private Sub R1C3_Click(sender As Object, e As EventArgs) Handles R1C3.Click

        CurrentRow = "1"
        CurrentColumn = "3"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R1C3.Text = CurrentNumber
            arrDone(2) = True
            arrValues(2) = CurrentNumber
        End If

    End Sub

    Private Sub R1C4_Click(sender As Object, e As EventArgs) Handles R1C4.Click

        CurrentRow = "1"
        CurrentColumn = "4"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R1C4.Text = CurrentNumber
            arrDone(3) = True
            arrValues(3) = CurrentNumber
        End If

    End Sub

    Private Sub R1C5_Click(sender As Object, e As EventArgs) Handles R1C5.Click

        CurrentRow = "1"
        CurrentColumn = "5"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R1C5.Text = CurrentNumber
            arrDone(4) = True
            arrValues(4) = CurrentNumber
        End If

    End Sub

    Private Sub R1C6_Click(sender As Object, e As EventArgs) Handles R1C6.Click

        CurrentRow = "1"
        CurrentColumn = "6"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R1C6.Text = CurrentNumber
            arrDone(5) = True
            arrValues(5) = CurrentNumber
        End If

    End Sub

    Private Sub R1C7_Click(sender As Object, e As EventArgs) Handles R1C7.Click

        CurrentRow = "1"
        CurrentColumn = "7"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R1C7.Text = CurrentNumber
            arrDone(6) = True
            arrValues(6) = CurrentNumber
        End If

    End Sub

    Private Sub R1C8_Click(sender As Object, e As EventArgs) Handles R1C8.Click

        CurrentRow = "1"
        CurrentColumn = "8"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R1C8.Text = CurrentNumber
            arrDone(7) = True
            arrValues(7) = CurrentNumber
        End If

    End Sub

    Private Sub R1C9_Click(sender As Object, e As EventArgs) Handles R1C9.Click

        CurrentRow = "1"
        CurrentColumn = "9"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R1C9.Text = CurrentNumber
            arrDone(8) = True
            arrValues(8) = CurrentNumber
        End If

    End Sub

    Private Sub R1C10_Click(sender As Object, e As EventArgs) Handles R1C10.Click

        CurrentRow = "1"
        CurrentColumn = "10"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R1C10.Text = CurrentNumber
            arrDone(9) = True
            arrValues(9) = CurrentNumber
        End If

    End Sub

    Private Sub R1C11_Click(sender As Object, e As EventArgs) Handles R1C11.Click

        CurrentRow = "1"
        CurrentColumn = "11"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R1C11.Text = CurrentNumber
            arrDone(10) = True
            arrValues(10) = CurrentNumber
        End If

    End Sub

    Private Sub R1C12_Click(sender As Object, e As EventArgs) Handles R1C12.Click

        CurrentRow = "1"
        CurrentColumn = "12"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R1C12.Text = CurrentNumber
            arrDone(11) = True
            arrValues(11) = CurrentNumber
        End If

    End Sub

    Private Sub R1C13_Click(sender As Object, e As EventArgs) Handles R1C13.Click

        CurrentRow = "1"
        CurrentColumn = "13"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R1C13.Text = CurrentNumber
            arrDone(12) = True
            arrValues(12) = CurrentNumber
        End If

    End Sub

    Private Sub R1C14_Click(sender As Object, e As EventArgs) Handles R1C14.Click

        CurrentRow = "1"
        CurrentColumn = "14"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R1C14.Text = CurrentNumber
            arrDone(13) = True
            arrValues(13) = CurrentNumber
        End If

    End Sub

    Private Sub R1C15_Click(sender As Object, e As EventArgs) Handles R1C15.Click

        CurrentRow = "1"
        CurrentColumn = "15"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R1C15.Text = CurrentNumber
            arrDone(14) = True
            arrValues(14) = CurrentNumber
        End If

    End Sub

    Private Sub R1C16_Click(sender As Object, e As EventArgs) Handles R1C16.Click

        CurrentRow = "1"
        CurrentColumn = "16"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R1C16.Text = CurrentNumber
            arrDone(15) = True
            arrValues(15) = CurrentNumber
        End If

    End Sub
    Private Sub R2C1_Click(sender As Object, e As EventArgs) Handles R2C1.Click

        CurrentRow = "2"
        CurrentColumn = "1"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R2C1.Text = CurrentNumber
            arrDone(16) = True
            arrValues(16) = CurrentNumber
        End If

    End Sub

    Private Sub R2C2_Click(sender As Object, e As EventArgs) Handles R2C2.Click

        CurrentRow = "2"
        CurrentColumn = "2"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R2C2.Text = CurrentNumber
            arrDone(17) = True
            arrValues(17) = CurrentNumber
        End If

    End Sub

    Private Sub R2C3_Click(sender As Object, e As EventArgs) Handles R2C3.Click

        CurrentRow = "2"
        CurrentColumn = "3"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R2C3.Text = CurrentNumber
            arrDone(18) = True
            arrValues(18) = CurrentNumber
        End If

    End Sub

    Private Sub R2C4_Click(sender As Object, e As EventArgs) Handles R2C4.Click
        CurrentRow = "2"
        CurrentColumn = "4"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R2C4.Text = CurrentNumber
            arrDone(19) = True
            arrValues(19) = CurrentNumber
        End If

    End Sub

    Private Sub R2C5_Click(sender As Object, e As EventArgs) Handles R2C5.Click

        CurrentRow = "2"
        CurrentColumn = "5"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R2C5.Text = CurrentNumber
            arrDone(20) = True
            arrValues(20) = CurrentNumber
        End If


    End Sub

    Private Sub R2C6_Click(sender As Object, e As EventArgs) Handles R2C6.Click

        CurrentRow = "2"
        CurrentColumn = "6"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R2C6.Text = CurrentNumber
            arrDone(21) = True
            arrValues(21) = CurrentNumber
        End If


    End Sub

    Private Sub R2C7_Click(sender As Object, e As EventArgs) Handles R2C7.Click

        CurrentRow = "2"
        CurrentColumn = "7"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R2C7.Text = CurrentNumber
            arrDone(22) = True
            arrValues(22) = CurrentNumber
        End If

    End Sub

    Private Sub R2C8_Click(sender As Object, e As EventArgs) Handles R2C8.Click

        CurrentRow = "2"
        CurrentColumn = "8"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R2C8.Text = CurrentNumber
            arrDone(23) = True
            arrValues(23) = CurrentNumber
        End If


    End Sub

    Private Sub R2C9_Click(sender As Object, e As EventArgs) Handles R2C9.Click

        CurrentRow = "2"
        CurrentColumn = "9"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R2C9.Text = CurrentNumber
            arrDone(24) = True
            arrValues(24) = CurrentNumber
        End If

    End Sub

    Private Sub R2C10_Click(sender As Object, e As EventArgs) Handles R2C10.Click

        CurrentRow = "2"
        CurrentColumn = "10"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R2C10.Text = CurrentNumber
            arrDone(25) = True
            arrValues(25) = CurrentNumber
        End If

    End Sub

    Private Sub R2C11_Click(sender As Object, e As EventArgs) Handles R2C11.Click

        CurrentRow = "2"
        CurrentColumn = "11"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R2C11.Text = CurrentNumber
            arrDone(26) = True
            arrValues(26) = CurrentNumber
        End If

    End Sub

    Private Sub R2C12_Click(sender As Object, e As EventArgs) Handles R2C12.Click

        CurrentRow = "2"
        CurrentColumn = "12"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R2C12.Text = CurrentNumber
            arrDone(27) = True
            arrValues(27) = CurrentNumber
        End If

    End Sub

    Private Sub R2C13_Click(sender As Object, e As EventArgs) Handles R2C13.Click

        CurrentRow = "2"
        CurrentColumn = "13"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R2C13.Text = CurrentNumber
            arrDone(28) = True
            arrValues(28) = CurrentNumber
        End If

    End Sub

    Private Sub R2C14_Click(sender As Object, e As EventArgs) Handles R2C14.Click

        CurrentRow = "2"
        CurrentColumn = "14"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R2C14.Text = CurrentNumber
            arrDone(29) = True
            arrValues(29) = CurrentNumber
        End If

    End Sub

    Private Sub R2C15_Click(sender As Object, e As EventArgs) Handles R2C15.Click

        CurrentRow = "2"
        CurrentColumn = "15"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R2C15.Text = CurrentNumber
            arrDone(30) = True
            arrValues(30) = CurrentNumber
        End If

    End Sub

    Private Sub R2C16_Click(sender As Object, e As EventArgs) Handles R2C16.Click

        CurrentRow = "2"
        CurrentColumn = "16"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R2C16.Text = CurrentNumber
            arrDone(31) = True
            arrValues(31) = CurrentNumber
        End If

    End Sub

    Private Sub R3C1_Click(sender As Object, e As EventArgs) Handles R3C1.Click

        CurrentRow = "3"
        CurrentColumn = "1"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R3C1.Text = CurrentNumber
            arrDone(32) = True
            arrValues(32) = CurrentNumber
        End If


    End Sub

    Private Sub R3C2_Click(sender As Object, e As EventArgs) Handles R3C2.Click

        CurrentRow = "3"
        CurrentColumn = "2"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R3C2.Text = CurrentNumber
            arrDone(33) = True
            arrValues(33) = CurrentNumber
        End If


    End Sub

    Private Sub R3C3_Click(sender As Object, e As EventArgs) Handles R3C3.Click

        CurrentRow = "3"
        CurrentColumn = "3"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R3C3.Text = CurrentNumber
            arrDone(34) = True
            arrValues(34) = CurrentNumber
        End If


    End Sub

    Private Sub R3C4_Click(sender As Object, e As EventArgs) Handles R3C4.Click

        CurrentRow = "3"
        CurrentColumn = "4"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R3C4.Text = CurrentNumber
            arrDone(35) = True
            arrValues(35) = CurrentNumber
        End If


    End Sub

    Private Sub R3C5_Click(sender As Object, e As EventArgs) Handles R3C5.Click

        CurrentRow = "3"
        CurrentColumn = "5"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R3C5.Text = CurrentNumber
            arrDone(36) = True
            arrValues(36) = CurrentNumber
        End If


    End Sub

    Private Sub R3C6_Click(sender As Object, e As EventArgs) Handles R3C6.Click

        CurrentRow = "3"
        CurrentColumn = "6"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R3C6.Text = CurrentNumber
            arrDone(37) = True
            arrValues(37) = CurrentNumber
        End If


    End Sub

    Private Sub R3C7_Click(sender As Object, e As EventArgs) Handles R3C7.Click

        CurrentRow = "3"
        CurrentColumn = "7"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R3C7.Text = CurrentNumber
            arrDone(38) = True
            arrValues(38) = CurrentNumber
        End If


    End Sub

    Private Sub R3C8_Click(sender As Object, e As EventArgs) Handles R3C8.Click

        CurrentRow = "3"
        CurrentColumn = "8"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R3C8.Text = CurrentNumber
            arrDone(39) = True
            arrValues(39) = CurrentNumber
        End If


    End Sub

    Private Sub R3C9_Click(sender As Object, e As EventArgs) Handles R3C9.Click

        CurrentRow = "3"
        CurrentColumn = "9"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R3C9.Text = CurrentNumber
            arrDone(40) = True
            arrValues(40) = CurrentNumber
        End If

    End Sub

    Private Sub R3C10_Click(sender As Object, e As EventArgs) Handles R3C10.Click

        CurrentRow = "3"
        CurrentColumn = "10"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R3C10.Text = CurrentNumber
            arrDone(41) = True
            arrValues(41) = CurrentNumber
        End If

    End Sub

    Private Sub R3C11_Click(sender As Object, e As EventArgs) Handles R3C11.Click

        CurrentRow = "3"
        CurrentColumn = "11"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R3C11.Text = CurrentNumber
            arrDone(42) = True
            arrValues(42) = CurrentNumber
        End If

    End Sub

    Private Sub R3C12_Click(sender As Object, e As EventArgs) Handles R3C12.Click

        CurrentRow = "3"
        CurrentColumn = "12"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R3C12.Text = CurrentNumber
            arrDone(43) = True
            arrValues(43) = CurrentNumber
        End If

    End Sub

    Private Sub R3C13_Click(sender As Object, e As EventArgs) Handles R3C13.Click

        CurrentRow = "3"
        CurrentColumn = "13"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R3C13.Text = CurrentNumber
            arrDone(44) = True
            arrValues(44) = CurrentNumber
        End If

    End Sub

    Private Sub R3C14_Click(sender As Object, e As EventArgs) Handles R3C14.Click

        CurrentRow = "3"
        CurrentColumn = "14"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R3C14.Text = CurrentNumber
            arrDone(45) = True
            arrValues(45) = CurrentNumber
        End If

    End Sub

    Private Sub R3C15_Click(sender As Object, e As EventArgs) Handles R3C15.Click

        CurrentRow = "3"
        CurrentColumn = "15"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R3C15.Text = CurrentNumber
            arrDone(46) = True
            arrValues(46) = CurrentNumber
        End If

    End Sub

    Private Sub R3C16_Click(sender As Object, e As EventArgs) Handles R3C16.Click

        CurrentRow = "3"
        CurrentColumn = "16"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R3C16.Text = CurrentNumber
            arrDone(47) = True
            arrValues(47) = CurrentNumber
        End If

    End Sub

    Private Sub R4C1_Click(sender As Object, e As EventArgs) Handles R4C1.Click

        CurrentRow = "4"
        CurrentColumn = "1"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R4C1.Text = CurrentNumber
            arrDone(48) = True
            arrValues(48) = CurrentNumber
        End If

    End Sub

    Private Sub R4C2_Click(sender As Object, e As EventArgs) Handles R4C2.Click

        CurrentRow = "4"
        CurrentColumn = "2"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R4C2.Text = CurrentNumber
            arrDone(49) = True
            arrValues(49) = CurrentNumber
        End If

    End Sub

    Private Sub R4C3_Click(sender As Object, e As EventArgs) Handles R4C3.Click

        CurrentRow = "4"
        CurrentColumn = "3"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R4C3.Text = CurrentNumber
            arrDone(50) = True
            arrValues(50) = CurrentNumber
        End If

    End Sub

    Private Sub R4C4_Click(sender As Object, e As EventArgs) Handles R4C4.Click

        CurrentRow = "4"
        CurrentColumn = "4"
        CurrentSquare = "1"

        If bRemoveANumberFromACell = False Then
            R4C4.Text = CurrentNumber
            arrDone(51) = True
            arrValues(51) = CurrentNumber
        End If

    End Sub

    Private Sub R4C5_Click(sender As Object, e As EventArgs) Handles R4C5.Click

        CurrentRow = "4"
        CurrentColumn = "5"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R4C5.Text = CurrentNumber
            arrDone(52) = True
            arrValues(52) = CurrentNumber
        End If

    End Sub

    Private Sub R4C6_Click(sender As Object, e As EventArgs) Handles R4C6.Click

        CurrentRow = "4"
        CurrentColumn = "6"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R4C6.Text = CurrentNumber
            arrDone(53) = True
            arrValues(53) = CurrentNumber
        End If

    End Sub

    Private Sub R4C7_Click(sender As Object, e As EventArgs) Handles R4C7.Click

        CurrentRow = "4"
        CurrentColumn = "7"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R4C7.Text = CurrentNumber
            arrDone(54) = True
            arrValues(54) = CurrentNumber
        End If

    End Sub

    Private Sub R4C8_Click(sender As Object, e As EventArgs) Handles R4C8.Click

        CurrentRow = "4"
        CurrentColumn = "8"
        CurrentSquare = "2"

        If bRemoveANumberFromACell = False Then
            R4C8.Text = CurrentNumber
            arrDone(55) = True
            arrValues(55) = CurrentNumber
        End If

    End Sub

    Private Sub R4C9_Click(sender As Object, e As EventArgs) Handles R4C9.Click

        CurrentRow = "4"
        CurrentColumn = "9"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R4C9.Text = CurrentNumber
            arrDone(56) = True
            arrValues(56) = CurrentNumber
        End If

    End Sub

    Private Sub R4C10_Click(sender As Object, e As EventArgs) Handles R4C10.Click

        CurrentRow = "4"
        CurrentColumn = "10"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R4C10.Text = CurrentNumber
            arrDone(57) = True
            arrValues(57) = CurrentNumber
        End If

    End Sub

    Private Sub R4C11_Click(sender As Object, e As EventArgs) Handles R4C11.Click

        CurrentRow = "4"
        CurrentColumn = "11"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R4C11.Text = CurrentNumber
            arrDone(58) = True
            arrValues(58) = CurrentNumber
        End If

    End Sub

    Private Sub R4C12_Click(sender As Object, e As EventArgs) Handles R4C12.Click

        CurrentRow = "4"
        CurrentColumn = "12"
        CurrentSquare = "3"

        If bRemoveANumberFromACell = False Then
            R4C12.Text = CurrentNumber
            arrDone(59) = True
            arrValues(59) = CurrentNumber
        End If

    End Sub

    Private Sub R4C13_Click(sender As Object, e As EventArgs) Handles R4C13.Click

        CurrentRow = "4"
        CurrentColumn = "13"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R4C13.Text = CurrentNumber
            arrDone(60) = True
            arrValues(60) = CurrentNumber
        End If

    End Sub

    Private Sub R4C14_Click(sender As Object, e As EventArgs) Handles R4C14.Click

        CurrentRow = "4"
        CurrentColumn = "14"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R4C14.Text = CurrentNumber
            arrDone(61) = True
            arrValues(61) = CurrentNumber
        End If

    End Sub

    Private Sub R4C15_Click(sender As Object, e As EventArgs) Handles R4C15.Click

        CurrentRow = "4"
        CurrentColumn = "15"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R4C15.Text = CurrentNumber
            arrDone(62) = True
            arrValues(62) = CurrentNumber
        End If

    End Sub

    Private Sub R4C16_Click(sender As Object, e As EventArgs) Handles R4C16.Click

        CurrentRow = "4"
        CurrentColumn = "16"
        CurrentSquare = "4"

        If bRemoveANumberFromACell = False Then
            R4C16.Text = CurrentNumber
            arrDone(63) = True
            arrValues(63) = CurrentNumber
        End If

    End Sub

    Private Sub R5C1_Click(sender As Object, e As EventArgs) Handles R5C1.Click

        CurrentRow = "5"
        CurrentColumn = "1"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R5C1.Text = CurrentNumber
            arrDone(64) = True
            arrValues(64) = CurrentNumber
        End If

    End Sub

    Private Sub R5C2_Click(sender As Object, e As EventArgs) Handles R5C2.Click

        CurrentRow = "5"
        CurrentColumn = "2"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R5C2.Text = CurrentNumber
            arrDone(65) = True
            arrValues(65) = CurrentNumber
        End If

    End Sub

    Private Sub R5C3_Click(sender As Object, e As EventArgs) Handles R5C3.Click

        CurrentRow = "5"
        CurrentColumn = "3"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R5C3.Text = CurrentNumber
            arrDone(66) = True
            arrValues(66) = CurrentNumber
        End If

    End Sub

    Private Sub R5C4_Click(sender As Object, e As EventArgs) Handles R5C4.Click

        CurrentRow = "5"
        CurrentColumn = "4"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R5C4.Text = CurrentNumber
            arrDone(67) = True
            arrValues(67) = CurrentNumber
        End If

    End Sub

    Private Sub R5C5_Click(sender As Object, e As EventArgs) Handles R5C5.Click

        CurrentRow = "5"
        CurrentColumn = "5"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R5C5.Text = CurrentNumber
            arrDone(68) = True
            arrValues(68) = CurrentNumber
        End If

    End Sub

    Private Sub R5C6_Click(sender As Object, e As EventArgs) Handles R5C6.Click

        CurrentRow = "5"
        CurrentColumn = "6"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R5C6.Text = CurrentNumber
            arrDone(69) = True
            arrValues(69) = CurrentNumber
        End If

    End Sub

    Private Sub R5C7_Click(sender As Object, e As EventArgs) Handles R5C7.Click

        CurrentRow = "5"
        CurrentColumn = "7"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R5C7.Text = CurrentNumber
            arrDone(70) = True
            arrValues(70) = CurrentNumber
        End If

    End Sub

    Private Sub R5C8_Click(sender As Object, e As EventArgs) Handles R5C8.Click

        CurrentRow = "5"
        CurrentColumn = "8"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R5C8.Text = CurrentNumber
            arrDone(71) = True
            arrValues(71) = CurrentNumber
        End If

    End Sub

    Private Sub R5C9_Click(sender As Object, e As EventArgs) Handles R5C9.Click

        CurrentRow = "5"
        CurrentColumn = "9"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R5C9.Text = CurrentNumber
            arrDone(72) = True
            arrValues(72) = CurrentNumber
        End If

    End Sub

    Private Sub R5C10_Click(sender As Object, e As EventArgs) Handles R5C10.Click

        CurrentRow = "5"
        CurrentColumn = "10"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R5C10.Text = CurrentNumber
            arrDone(73) = True
            arrValues(73) = CurrentNumber
        End If

    End Sub

    Private Sub R5C11_Click(sender As Object, e As EventArgs) Handles R5C11.Click

        CurrentRow = "5"
        CurrentColumn = "11"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R5C11.Text = CurrentNumber
            arrDone(74) = True
            arrValues(74) = CurrentNumber
        End If

    End Sub

    Private Sub R5C12_Click(sender As Object, e As EventArgs) Handles R5C12.Click

        CurrentRow = "5"
        CurrentColumn = "12"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R5C12.Text = CurrentNumber
            arrDone(75) = True
            arrValues(75) = CurrentNumber
        End If

    End Sub

    Private Sub R5C13_Click(sender As Object, e As EventArgs) Handles R5C13.Click

        CurrentRow = "5"
        CurrentColumn = "13"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R5C13.Text = CurrentNumber
            arrDone(76) = True
            arrValues(76) = CurrentNumber
        End If

    End Sub

    Private Sub R5C14_Click(sender As Object, e As EventArgs) Handles R5C14.Click

        CurrentRow = "5"
        CurrentColumn = "14"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R5C14.Text = CurrentNumber
            arrDone(77) = True
            arrValues(77) = CurrentNumber
        End If

    End Sub

    Private Sub R5C15_Click(sender As Object, e As EventArgs) Handles R5C15.Click

        CurrentRow = "5"
        CurrentColumn = "15"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R5C15.Text = CurrentNumber
            arrDone(78) = True
            arrValues(78) = CurrentNumber
        End If

    End Sub

    Private Sub R5C16_Click(sender As Object, e As EventArgs) Handles R5C16.Click

        CurrentRow = "5"
        CurrentColumn = "16"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R5C16.Text = CurrentNumber
            arrDone(79) = True
            arrValues(79) = CurrentNumber
        End If

    End Sub

    Private Sub R6C1_Click(sender As Object, e As EventArgs) Handles R6C1.Click

        CurrentRow = "6"
        CurrentColumn = "1"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R6C1.Text = CurrentNumber
            arrDone(80) = True
            arrValues(80) = CurrentNumber
        End If

    End Sub

    Private Sub R6C2_Click(sender As Object, e As EventArgs) Handles R6C2.Click

        CurrentRow = "6"
        CurrentColumn = "2"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R6C2.Text = CurrentNumber
            arrDone(81) = True
            arrValues(81) = CurrentNumber
        End If

    End Sub

    Private Sub R6C3_Click(sender As Object, e As EventArgs) Handles R6C3.Click

        CurrentRow = "6"
        CurrentColumn = "3"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R6C3.Text = CurrentNumber
            arrDone(82) = True
            arrValues(82) = CurrentNumber
        End If

    End Sub

    Private Sub R6C4_Click(sender As Object, e As EventArgs) Handles R6C4.Click

        CurrentRow = "6"
        CurrentColumn = "4"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R6C4.Text = CurrentNumber
            arrDone(83) = True
            arrValues(83) = CurrentNumber
        End If

    End Sub

    Private Sub R6C5_Click(sender As Object, e As EventArgs) Handles R6C5.Click

        CurrentRow = "6"
        CurrentColumn = "5"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R6C5.Text = CurrentNumber
            arrDone(84) = True
            arrValues(84) = CurrentNumber
        End If

    End Sub

    Private Sub R6C6_Click(sender As Object, e As EventArgs) Handles R6C6.Click

        CurrentRow = "6"
        CurrentColumn = "6"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R6C6.Text = CurrentNumber
            arrDone(85) = True
            arrValues(85) = CurrentNumber
        End If

    End Sub

    Private Sub R6C7_Click(sender As Object, e As EventArgs) Handles R6C7.Click

        CurrentRow = "6"
        CurrentColumn = "7"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R6C7.Text = CurrentNumber
            arrDone(86) = True
            arrValues(86) = CurrentNumber
        End If

    End Sub

    Private Sub R6C8_Click(sender As Object, e As EventArgs) Handles R6C8.Click

        CurrentRow = "6"
        CurrentColumn = "8"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R6C8.Text = CurrentNumber
            arrDone(87) = True
            arrValues(87) = CurrentNumber
        End If

    End Sub

    Private Sub R6C9_Click(sender As Object, e As EventArgs) Handles R6C9.Click

        CurrentRow = "6"
        CurrentColumn = "9"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R6C9.Text = CurrentNumber
            arrDone(88) = True
            arrValues(88) = CurrentNumber
        End If

    End Sub

    Private Sub R6C10_Click(sender As Object, e As EventArgs) Handles R6C10.Click

        CurrentRow = "6"
        CurrentColumn = "10"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R6C10.Text = CurrentNumber
            arrDone(89) = True
            arrValues(89) = CurrentNumber
        End If

    End Sub

    Private Sub R6C11_Click(sender As Object, e As EventArgs) Handles R6C11.Click

        CurrentRow = "6"
        CurrentColumn = "11"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R6C11.Text = CurrentNumber
            arrDone(90) = True
            arrValues(90) = CurrentNumber
        End If

    End Sub

    Private Sub R6C12_Click(sender As Object, e As EventArgs) Handles R6C12.Click

        CurrentRow = "6"
        CurrentColumn = "12"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R6C12.Text = CurrentNumber
            arrDone(91) = True
            arrValues(91) = CurrentNumber
        End If

    End Sub

    Private Sub R6C13_Click(sender As Object, e As EventArgs) Handles R6C13.Click

        CurrentRow = "6"
        CurrentColumn = "13"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R6C13.Text = CurrentNumber
            arrDone(92) = True
            arrValues(92) = CurrentNumber
        End If

    End Sub
    Private Sub R6C14_Click(sender As Object, e As EventArgs) Handles R6C14.Click

        CurrentRow = "6"
        CurrentColumn = "14"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R6C14.Text = CurrentNumber
            arrDone(93) = True
            arrValues(93) = CurrentNumber
        End If

    End Sub
    Private Sub R6C15_Click(sender As Object, e As EventArgs) Handles R6C15.Click

        CurrentRow = "6"
        CurrentColumn = "15"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R6C15.Text = CurrentNumber
            arrDone(94) = True
            arrValues(94) = CurrentNumber
        End If

    End Sub
    Private Sub R6C16_Click(sender As Object, e As EventArgs) Handles R6C16.Click

        CurrentRow = "6"
        CurrentColumn = "16"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R6C16.Text = CurrentNumber
            arrDone(95) = True
            arrValues(95) = CurrentNumber
        End If

    End Sub

    Private Sub R7C1_Click(sender As Object, e As EventArgs) Handles R7C1.Click

        CurrentRow = "7"
        CurrentColumn = "1"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R7C1.Text = CurrentNumber
            arrDone(96) = True
            arrValues(96) = CurrentNumber
        End If

    End Sub

    Private Sub R7C2_Click(sender As Object, e As EventArgs) Handles R7C2.Click

        CurrentRow = "7"
        CurrentColumn = "2"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R7C2.Text = CurrentNumber
            arrDone(97) = True
            arrValues(97) = CurrentNumber
        End If

    End Sub

    Private Sub R7C3_Click(sender As Object, e As EventArgs) Handles R7C3.Click

        CurrentRow = "7"
        CurrentColumn = "3"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R7C3.Text = CurrentNumber
            arrDone(98) = True
            arrValues(98) = CurrentNumber
        End If

    End Sub

    Private Sub R7C4_Click(sender As Object, e As EventArgs) Handles R7C4.Click

        CurrentRow = "7"
        CurrentColumn = "4"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R7C4.Text = CurrentNumber
            arrDone(99) = True
            arrValues(99) = CurrentNumber
        End If

    End Sub

    Private Sub R7C5_Click(sender As Object, e As EventArgs) Handles R7C5.Click

        CurrentRow = "7"
        CurrentColumn = "5"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R7C5.Text = CurrentNumber
            arrDone(100) = True
            arrValues(100) = CurrentNumber
        End If

    End Sub

    Private Sub R7C6_Click(sender As Object, e As EventArgs) Handles R7C6.Click

        CurrentRow = "7"
        CurrentColumn = "6"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R7C6.Text = CurrentNumber
            arrDone(101) = True
            arrValues(101) = CurrentNumber
        End If

    End Sub

    Private Sub R7C7_Click(sender As Object, e As EventArgs) Handles R7C7.Click

        CurrentRow = "7"
        CurrentColumn = "7"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R7C7.Text = CurrentNumber
            arrDone(102) = True
            arrValues(102) = CurrentNumber
        End If

    End Sub

    Private Sub R7C8_Click(sender As Object, e As EventArgs) Handles R7C8.Click

        CurrentRow = "7"
        CurrentColumn = "8"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R7C8.Text = CurrentNumber
            arrDone(103) = True
            arrValues(103) = CurrentNumber
        End If

    End Sub

    Private Sub R7C9_Click(sender As Object, e As EventArgs) Handles R7C9.Click

        CurrentRow = "7"
        CurrentColumn = "9"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R7C9.Text = CurrentNumber
            arrDone(104) = True
            arrValues(104) = CurrentNumber
        End If

    End Sub

    Private Sub R7C10_Click(sender As Object, e As EventArgs) Handles R7C10.Click

        CurrentRow = "7"
        CurrentColumn = "10"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R7C10.Text = CurrentNumber
            arrDone(105) = True
            arrValues(105) = CurrentNumber
        End If

    End Sub

    Private Sub R7C11_Click(sender As Object, e As EventArgs) Handles R7C11.Click

        CurrentRow = "7"
        CurrentColumn = "11"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R7C11.Text = CurrentNumber
            arrDone(106) = True
            arrValues(106) = CurrentNumber
        End If

    End Sub

    Private Sub R7C12_Click(sender As Object, e As EventArgs) Handles R7C12.Click

        CurrentRow = "7"
        CurrentColumn = "12"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R7C12.Text = CurrentNumber
            arrDone(107) = True
            arrValues(107) = CurrentNumber
        End If

    End Sub

    Private Sub R7C13_Click(sender As Object, e As EventArgs) Handles R7C13.Click

        CurrentRow = "7"
        CurrentColumn = "13"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R7C13.Text = CurrentNumber
            arrDone(108) = True
            arrValues(108) = CurrentNumber
        End If

    End Sub

    Private Sub R7C14_Click(sender As Object, e As EventArgs) Handles R7C14.Click

        CurrentRow = "7"
        CurrentColumn = "14"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R7C14.Text = CurrentNumber
            arrDone(109) = True
            arrValues(109) = CurrentNumber
        End If

    End Sub

    Private Sub R7C15_Click(sender As Object, e As EventArgs) Handles R7C15.Click

        CurrentRow = "7"
        CurrentColumn = "15"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R7C15.Text = CurrentNumber
            arrDone(110) = True
            arrValues(110) = CurrentNumber
        End If

    End Sub

    Private Sub R7C16_Click(sender As Object, e As EventArgs) Handles R7C16.Click

        CurrentRow = "7"
        CurrentColumn = "16"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R7C16.Text = CurrentNumber
            arrDone(111) = True
            arrValues(111) = CurrentNumber
        End If

    End Sub

    Private Sub R8C1_Click(sender As Object, e As EventArgs) Handles R8C1.Click

        CurrentRow = "8"
        CurrentColumn = "1"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R8C1.Text = CurrentNumber
            arrDone(112) = True
            arrValues(112) = CurrentNumber
        End If

    End Sub

    Private Sub R8C2_Click(sender As Object, e As EventArgs) Handles R8C2.Click

        CurrentRow = "8"
        CurrentColumn = "2"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R8C2.Text = CurrentNumber
            arrDone(113) = True
            arrValues(113) = CurrentNumber
        End If

    End Sub

    Private Sub R8C3_Click(sender As Object, e As EventArgs) Handles R8C3.Click

        CurrentRow = "8"
        CurrentColumn = "3"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R8C3.Text = CurrentNumber
            arrDone(114) = True
            arrValues(114) = CurrentNumber
        End If

    End Sub

    Private Sub R8C4_Click(sender As Object, e As EventArgs) Handles R8C4.Click

        CurrentRow = "8"
        CurrentColumn = "4"
        CurrentSquare = "5"

        If bRemoveANumberFromACell = False Then
            R8C4.Text = CurrentNumber
            arrDone(115) = True
            arrValues(115) = CurrentNumber
        End If

    End Sub

    Private Sub R8C5_Click(sender As Object, e As EventArgs) Handles R8C5.Click

        CurrentRow = "8"
        CurrentColumn = "5"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R8C5.Text = CurrentNumber
            arrDone(116) = True
            arrValues(116) = CurrentNumber
        End If

    End Sub

    Private Sub R8C6_Click(sender As Object, e As EventArgs) Handles R8C6.Click

        CurrentRow = "8"
        CurrentColumn = "6"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R8C6.Text = CurrentNumber
            arrDone(117) = True
            arrValues(117) = CurrentNumber
        End If

    End Sub

    Private Sub R8C7_Click(sender As Object, e As EventArgs) Handles R8C7.Click

        CurrentRow = "8"
        CurrentColumn = "7"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R8C7.Text = CurrentNumber
            arrDone(118) = True
            arrValues(118) = CurrentNumber
        End If

    End Sub

    Private Sub R8C8_Click(sender As Object, e As EventArgs) Handles R8C8.Click

        CurrentRow = "8"
        CurrentColumn = "8"
        CurrentSquare = "6"

        If bRemoveANumberFromACell = False Then
            R8C8.Text = CurrentNumber
            arrDone(119) = True
            arrValues(119) = CurrentNumber
        End If

    End Sub

    Private Sub R8C9_Click(sender As Object, e As EventArgs) Handles R8C9.Click

        CurrentRow = "8"
        CurrentColumn = "9"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R8C9.Text = CurrentNumber
            arrDone(120) = True
            arrValues(120) = CurrentNumber
        End If

    End Sub

    Private Sub R8C10_Click(sender As Object, e As EventArgs) Handles R8C10.Click

        CurrentRow = "8"
        CurrentColumn = "10"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R8C10.Text = CurrentNumber
            arrDone(121) = True
            arrValues(121) = CurrentNumber
        End If

    End Sub

    Private Sub R8C11_Click(sender As Object, e As EventArgs) Handles R8C11.Click

        CurrentRow = "8"
        CurrentColumn = "11"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R8C11.Text = CurrentNumber
            arrDone(122) = True
            arrValues(122) = CurrentNumber
        End If

    End Sub

    Private Sub R8C12_Click(sender As Object, e As EventArgs) Handles R8C12.Click

        CurrentRow = "8"
        CurrentColumn = "12"
        CurrentSquare = "7"

        If bRemoveANumberFromACell = False Then
            R8C12.Text = CurrentNumber
            arrDone(123) = True
            arrValues(123) = CurrentNumber
        End If

    End Sub

    Private Sub R8C13_Click(sender As Object, e As EventArgs) Handles R8C13.Click

        CurrentRow = "8"
        CurrentColumn = "13"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R8C13.Text = CurrentNumber
            arrDone(124) = True
            arrValues(124) = CurrentNumber
        End If

    End Sub

    Private Sub R8C14_Click(sender As Object, e As EventArgs) Handles R8C14.Click

        CurrentRow = "8"
        CurrentColumn = "14"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R8C14.Text = CurrentNumber
            arrDone(125) = True
            arrValues(125) = CurrentNumber
        End If

    End Sub

    Private Sub R8C15_Click(sender As Object, e As EventArgs) Handles R8C15.Click

        CurrentRow = "8"
        CurrentColumn = "15"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R8C15.Text = CurrentNumber
            arrDone(126) = True
            arrValues(126) = CurrentNumber
        End If

    End Sub

    Private Sub R8C16_Click(sender As Object, e As EventArgs) Handles R8C16.Click

        CurrentRow = "8"
        CurrentColumn = "16"
        CurrentSquare = "8"

        If bRemoveANumberFromACell = False Then
            R8C16.Text = CurrentNumber
            arrDone(127) = True
            arrValues(127) = CurrentNumber
        End If

    End Sub

    Private Sub R9C1_Click(sender As Object, e As EventArgs) Handles R9C1.Click

        CurrentRow = "9"
        CurrentColumn = "1"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R9C1.Text = CurrentNumber
            arrDone(128) = True
            arrValues(128) = CurrentNumber
        End If

    End Sub

    Private Sub R9C2_Click(sender As Object, e As EventArgs) Handles R9C2.Click

        CurrentRow = "9"
        CurrentColumn = "2"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R9C2.Text = CurrentNumber
            arrDone(129) = True
            arrValues(129) = CurrentNumber
        End If

    End Sub

    Private Sub R9C3_Click(sender As Object, e As EventArgs) Handles R9C3.Click

        CurrentRow = "9"
        CurrentColumn = "3"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R9C3.Text = CurrentNumber
            arrDone(130) = True
            arrValues(130) = CurrentNumber
        End If

    End Sub

    Private Sub R9C4_Click(sender As Object, e As EventArgs) Handles R9C4.Click

        CurrentRow = "9"
        CurrentColumn = "4"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R9C4.Text = CurrentNumber
            arrDone(131) = True
            arrValues(131) = CurrentNumber
        End If

    End Sub

    Private Sub R9C5_Click(sender As Object, e As EventArgs) Handles R9C5.Click

        CurrentRow = "9"
        CurrentColumn = "5"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R9C5.Text = CurrentNumber
            arrDone(132) = True
            arrValues(132) = CurrentNumber
        End If

    End Sub

    Private Sub R9C6_Click(sender As Object, e As EventArgs) Handles R9C6.Click

        CurrentRow = "9"
        CurrentColumn = "6"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R9C6.Text = CurrentNumber
            arrDone(133) = True
            arrValues(133) = CurrentNumber
        End If

    End Sub

    Private Sub R9C7_Click(sender As Object, e As EventArgs) Handles R9C7.Click

        CurrentRow = "9"
        CurrentColumn = "7"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R9C7.Text = CurrentNumber
            arrDone(134) = True
            arrValues(134) = CurrentNumber
        End If

    End Sub

    Private Sub R9C8_Click(sender As Object, e As EventArgs) Handles R9C8.Click

        CurrentRow = "9"
        CurrentColumn = "8"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R9C8.Text = CurrentNumber
            arrDone(135) = True
            arrValues(135) = CurrentNumber
        End If

    End Sub

    Private Sub R9C9_Click(sender As Object, e As EventArgs) Handles R9C9.Click

        CurrentRow = "9"
        CurrentColumn = "9"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R9C9.Text = CurrentNumber
            arrDone(136) = True
            arrValues(136) = CurrentNumber
        End If

    End Sub

    Private Sub R9C10_Click(sender As Object, e As EventArgs) Handles R9C10.Click

        CurrentRow = "9"
        CurrentColumn = "10"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R9C10.Text = CurrentNumber
            arrDone(137) = True
            arrValues(137) = CurrentNumber
        End If

    End Sub

    Private Sub R9C11_Click(sender As Object, e As EventArgs) Handles R9C11.Click

        CurrentRow = "9"
        CurrentColumn = "11"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R9C11.Text = CurrentNumber
            arrDone(138) = True
            arrValues(138) = CurrentNumber
        End If

    End Sub

    Private Sub R9C12_Click(sender As Object, e As EventArgs) Handles R9C12.Click

        CurrentRow = "9"
        CurrentColumn = "12"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R9C12.Text = CurrentNumber
            arrDone(139) = True
            arrValues(139) = CurrentNumber
        End If

    End Sub

    Private Sub R9C13_Click(sender As Object, e As EventArgs) Handles R9C13.Click

        CurrentRow = "9"
        CurrentColumn = "13"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R9C13.Text = CurrentNumber
            arrDone(140) = True
            arrValues(140) = CurrentNumber
        End If

    End Sub

    Private Sub R9C14_Click(sender As Object, e As EventArgs) Handles R9C14.Click

        CurrentRow = "9"
        CurrentColumn = "14"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R9C14.Text = CurrentNumber
            arrDone(141) = True
            arrValues(141) = CurrentNumber
        End If

    End Sub

    Private Sub R9C15_Click(sender As Object, e As EventArgs) Handles R9C15.Click

        CurrentRow = "9"
        CurrentColumn = "15"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R9C15.Text = CurrentNumber
            arrDone(142) = True
            arrValues(142) = CurrentNumber
        End If

    End Sub

    Private Sub R9C16_Click(sender As Object, e As EventArgs) Handles R9C16.Click

        CurrentRow = "9"
        CurrentColumn = "16"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R9C16.Text = CurrentNumber
            arrDone(143) = True
            arrValues(143) = CurrentNumber
        End If

    End Sub

    Private Sub R10C1_Click(sender As Object, e As EventArgs) Handles R10C1.Click

        CurrentRow = "10"
        CurrentColumn = "1"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R10C1.Text = CurrentNumber
            arrDone(144) = True
            arrValues(144) = CurrentNumber
        End If

    End Sub

    Private Sub R10C2_Click(sender As Object, e As EventArgs) Handles R10C2.Click

        CurrentRow = "10"
        CurrentColumn = "2"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R10C2.Text = CurrentNumber
            arrDone(145) = True
            arrValues(145) = CurrentNumber
        End If

    End Sub

    Private Sub R10C3_Click(sender As Object, e As EventArgs) Handles R10C3.Click

        CurrentRow = "10"
        CurrentColumn = "3"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R10C3.Text = CurrentNumber
            arrDone(146) = True
            arrValues(146) = CurrentNumber
        End If

    End Sub

    Private Sub R10C4_Click(sender As Object, e As EventArgs) Handles R10C4.Click

        CurrentRow = "10"
        CurrentColumn = "4"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R10C4.Text = CurrentNumber
            arrDone(147) = True
            arrValues(147) = CurrentNumber
        End If

    End Sub

    Private Sub R10C5_Click(sender As Object, e As EventArgs) Handles R10C5.Click

        CurrentRow = "10"
        CurrentColumn = "5"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R10C5.Text = CurrentNumber
            arrDone(148) = True
            arrValues(148) = CurrentNumber
        End If

    End Sub

    Private Sub R10C6_Click(sender As Object, e As EventArgs) Handles R10C6.Click

        CurrentRow = "10"
        CurrentColumn = "6"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R10C6.Text = CurrentNumber
            arrDone(149) = True
            arrValues(149) = CurrentNumber
        End If

    End Sub

    Private Sub R10C7_Click(sender As Object, e As EventArgs) Handles R10C7.Click

        CurrentRow = "10"
        CurrentColumn = "7"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R10C7.Text = CurrentNumber
            arrDone(150) = True
            arrValues(150) = CurrentNumber
        End If

    End Sub

    Private Sub R10C8_Click(sender As Object, e As EventArgs) Handles R10C8.Click

        CurrentRow = "10"
        CurrentColumn = "8"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R10C8.Text = CurrentNumber
            arrDone(151) = True
            arrValues(151) = CurrentNumber
        End If

    End Sub

    Private Sub R10C9_Click(sender As Object, e As EventArgs) Handles R10C9.Click

        CurrentRow = "10"
        CurrentColumn = "9"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R10C9.Text = CurrentNumber
            arrDone(152) = True
            arrValues(152) = CurrentNumber
        End If

    End Sub

    Private Sub R10C10_Click(sender As Object, e As EventArgs) Handles R10C10.Click

        CurrentRow = "10"
        CurrentColumn = "10"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R10C10.Text = CurrentNumber
            arrDone(153) = True
            arrValues(153) = CurrentNumber
        End If

    End Sub

    Private Sub R10C11_Click(sender As Object, e As EventArgs) Handles R10C11.Click

        CurrentRow = "10"
        CurrentColumn = "11"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R10C11.Text = CurrentNumber
            arrDone(154) = True
            arrValues(154) = CurrentNumber
        End If

    End Sub

    Private Sub R10C12_Click(sender As Object, e As EventArgs) Handles R10C12.Click

        CurrentRow = "10"
        CurrentColumn = "12"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R10C12.Text = CurrentNumber
            arrDone(155) = True
            arrValues(155) = CurrentNumber
        End If

    End Sub

    Private Sub R10C13_Click(sender As Object, e As EventArgs) Handles R10C13.Click

        CurrentRow = "10"
        CurrentColumn = "13"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R10C13.Text = CurrentNumber
            arrDone(156) = True
            arrValues(156) = CurrentNumber
        End If

    End Sub

    Private Sub R10C14_Click(sender As Object, e As EventArgs) Handles R10C14.Click

        CurrentRow = "10"
        CurrentColumn = "14"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R10C14.Text = CurrentNumber
            arrDone(157) = True
            arrValues(157) = CurrentNumber
        End If

    End Sub

    Private Sub R10C15_Click(sender As Object, e As EventArgs) Handles R10C15.Click

        CurrentRow = "10"
        CurrentColumn = "15"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R10C15.Text = CurrentNumber
            arrDone(158) = True
            arrValues(158) = CurrentNumber
        End If

    End Sub

    Private Sub R10C16_Click(sender As Object, e As EventArgs) Handles R10C16.Click

        CurrentRow = "10"
        CurrentColumn = "16"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R10C16.Text = CurrentNumber
            arrDone(159) = True
            arrValues(159) = CurrentNumber
        End If

    End Sub

    Private Sub R11C1_Click(sender As Object, e As EventArgs) Handles R11C1.Click

        CurrentRow = "11"
        CurrentColumn = "1"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R11C1.Text = CurrentNumber
            arrDone(160) = True
            arrValues(160) = CurrentNumber
        End If

    End Sub

    Private Sub R11C2_Click(sender As Object, e As EventArgs) Handles R11C2.Click

        CurrentRow = "11"
        CurrentColumn = "2"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R11C2.Text = CurrentNumber
            arrDone(161) = True
            arrValues(161) = CurrentNumber
        End If

    End Sub

    Private Sub R11C3_Click(sender As Object, e As EventArgs) Handles R11C3.Click

        CurrentRow = "11"
        CurrentColumn = "3"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R11C3.Text = CurrentNumber
            arrDone(162) = True
            arrValues(162) = CurrentNumber
        End If

    End Sub

    Private Sub R11C4_Click(sender As Object, e As EventArgs) Handles R11C4.Click

        CurrentRow = "11"
        CurrentColumn = "4"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R11C4.Text = CurrentNumber
            arrDone(163) = True
            arrValues(163) = CurrentNumber
        End If

    End Sub

    Private Sub R11C5_Click(sender As Object, e As EventArgs) Handles R11C5.Click

        CurrentRow = "11"
        CurrentColumn = "5"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R11C5.Text = CurrentNumber
            arrDone(164) = True
            arrValues(164) = CurrentNumber
        End If

    End Sub

    Private Sub R11C6_Click(sender As Object, e As EventArgs) Handles R11C6.Click

        CurrentRow = "11"
        CurrentColumn = "6"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R11C6.Text = CurrentNumber
            arrDone(165) = True
            arrValues(165) = CurrentNumber
        End If

    End Sub

    Private Sub R11C7_Click(sender As Object, e As EventArgs) Handles R11C7.Click

        CurrentRow = "11"
        CurrentColumn = "7"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R11C7.Text = CurrentNumber
            arrDone(166) = True
            arrValues(166) = CurrentNumber
        End If

    End Sub

    Private Sub R11C8_Click(sender As Object, e As EventArgs) Handles R11C8.Click

        CurrentRow = "11"
        CurrentColumn = "8"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R11C8.Text = CurrentNumber
            arrDone(167) = True
            arrValues(167) = CurrentNumber
        End If

    End Sub

    Private Sub R11C9_Click(sender As Object, e As EventArgs) Handles R11C9.Click

        CurrentRow = "11"
        CurrentColumn = "9"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R11C9.Text = CurrentNumber
            arrDone(168) = True
            arrValues(168) = CurrentNumber
        End If

    End Sub

    Private Sub R11C10_Click(sender As Object, e As EventArgs) Handles R11C10.Click

        CurrentRow = "11"
        CurrentColumn = "10"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R11C10.Text = CurrentNumber
            arrDone(169) = True
            arrValues(169) = CurrentNumber
        End If

    End Sub

    Private Sub R11C11_Click(sender As Object, e As EventArgs) Handles R11C11.Click

        CurrentRow = "11"
        CurrentColumn = "11"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R11C11.Text = CurrentNumber
            arrDone(170) = True
            arrValues(170) = CurrentNumber
        End If

    End Sub

    Private Sub R11C12_Click(sender As Object, e As EventArgs) Handles R11C12.Click

        CurrentRow = "11"
        CurrentColumn = "12"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R11C12.Text = CurrentNumber
            arrDone(171) = True
            arrValues(171) = CurrentNumber
        End If

    End Sub

    Private Sub R11C13_Click(sender As Object, e As EventArgs) Handles R11C13.Click

        CurrentRow = "11"
        CurrentColumn = "13"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R11C13.Text = CurrentNumber
            arrDone(172) = True
            arrValues(172) = CurrentNumber
        End If

    End Sub

    Private Sub R11C14_Click(sender As Object, e As EventArgs) Handles R11C14.Click

        CurrentRow = "11"
        CurrentColumn = "14"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R11C14.Text = CurrentNumber
            arrDone(173) = True
            arrValues(173) = CurrentNumber
        End If

    End Sub

    Private Sub R11C15_Click(sender As Object, e As EventArgs) Handles R11C15.Click

        CurrentRow = "11"
        CurrentColumn = "15"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R11C15.Text = CurrentNumber
            arrDone(174) = True
            arrValues(174) = CurrentNumber
        End If

    End Sub

    Private Sub R11C16_Click(sender As Object, e As EventArgs) Handles R11C16.Click

        CurrentRow = "11"
        CurrentColumn = "16"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R11C16.Text = CurrentNumber
            arrDone(175) = True
            arrValues(175) = CurrentNumber
        End If

    End Sub
    '*****

    Private Sub R12C1_Click(sender As Object, e As EventArgs) Handles R12C1.Click

        CurrentRow = "12"
        CurrentColumn = "1"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R12C1.Text = CurrentNumber
            arrDone(176) = True
            arrValues(176) = CurrentNumber
        End If

    End Sub

    Private Sub R12C2_Click(sender As Object, e As EventArgs) Handles R12C2.Click

        CurrentRow = "12"
        CurrentColumn = "2"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R12C2.Text = CurrentNumber
            arrDone(177) = True
            arrValues(177) = CurrentNumber
        End If

    End Sub

    Private Sub R12C3_Click(sender As Object, e As EventArgs) Handles R12C3.Click

        CurrentRow = "12"
        CurrentColumn = "3"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R12C3.Text = CurrentNumber
            arrDone(178) = True
            arrValues(178) = CurrentNumber
        End If

    End Sub

    Private Sub R12C4_Click(sender As Object, e As EventArgs) Handles R12C4.Click

        CurrentRow = "12"
        CurrentColumn = "4"
        CurrentSquare = "9"

        If bRemoveANumberFromACell = False Then
            R12C4.Text = CurrentNumber
            arrDone(179) = True
            arrValues(179) = CurrentNumber
        End If

    End Sub

    Private Sub R12C5_Click(sender As Object, e As EventArgs) Handles R12C5.Click

        CurrentRow = "12"
        CurrentColumn = "5"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R12C5.Text = CurrentNumber
            arrDone(180) = True
            arrValues(180) = CurrentNumber
        End If

    End Sub

    Private Sub R12C6_Click(sender As Object, e As EventArgs) Handles R12C6.Click

        CurrentRow = "12"
        CurrentColumn = "6"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R12C6.Text = CurrentNumber
            arrDone(181) = True
            arrValues(181) = CurrentNumber
        End If

    End Sub

    Private Sub R12C7_Click(sender As Object, e As EventArgs) Handles R12C7.Click

        CurrentRow = "12"
        CurrentColumn = "7"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R12C7.Text = CurrentNumber
            arrDone(182) = True
            arrValues(182) = CurrentNumber
        End If

    End Sub

    Private Sub R12C8_Click(sender As Object, e As EventArgs) Handles R12C8.Click

        CurrentRow = "12"
        CurrentColumn = "8"
        CurrentSquare = "10"

        If bRemoveANumberFromACell = False Then
            R12C8.Text = CurrentNumber
            arrDone(183) = True
            arrValues(183) = CurrentNumber
        End If

    End Sub

    Private Sub R12C9_Click(sender As Object, e As EventArgs) Handles R12C9.Click

        CurrentRow = "12"
        CurrentColumn = "9"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R12C9.Text = CurrentNumber
            arrDone(184) = True
            arrValues(184) = CurrentNumber
        End If

    End Sub

    Private Sub R12C10_Click(sender As Object, e As EventArgs) Handles R12C10.Click

        CurrentRow = "12"
        CurrentColumn = "10"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R12C10.Text = CurrentNumber
            arrDone(185) = True
            arrValues(185) = CurrentNumber
        End If

    End Sub

    Private Sub R12C11_Click(sender As Object, e As EventArgs) Handles R12C11.Click

        CurrentRow = "12"
        CurrentColumn = "11"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R12C11.Text = CurrentNumber
            arrDone(186) = True
            arrValues(186) = CurrentNumber
        End If

    End Sub

    Private Sub R12C12_Click(sender As Object, e As EventArgs) Handles R12C12.Click

        CurrentRow = "12"
        CurrentColumn = "12"
        CurrentSquare = "11"

        If bRemoveANumberFromACell = False Then
            R12C12.Text = CurrentNumber
            arrDone(187) = True
            arrValues(187) = CurrentNumber
        End If

    End Sub

    Private Sub R12C13_Click(sender As Object, e As EventArgs) Handles R12C13.Click

        CurrentRow = "12"
        CurrentColumn = "13"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R12C13.Text = CurrentNumber
            arrDone(188) = True
            arrValues(188) = CurrentNumber
        End If

    End Sub

    Private Sub R12C14_Click(sender As Object, e As EventArgs) Handles R12C14.Click

        CurrentRow = "12"
        CurrentColumn = "14"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R12C14.Text = CurrentNumber
            arrDone(189) = True
            arrValues(189) = CurrentNumber
        End If

    End Sub

    Private Sub R12C15_Click(sender As Object, e As EventArgs) Handles R12C15.Click

        CurrentRow = "12"
        CurrentColumn = "15"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R12C15.Text = CurrentNumber
            arrDone(190) = True
            arrValues(190) = CurrentNumber
        End If

    End Sub

    Private Sub R12C16_Click(sender As Object, e As EventArgs) Handles R12C16.Click

        CurrentRow = "12"
        CurrentColumn = "16"
        CurrentSquare = "12"

        If bRemoveANumberFromACell = False Then
            R12C16.Text = CurrentNumber
            arrDone(191) = True
            arrValues(191) = CurrentNumber
        End If

    End Sub

    Private Sub R13C1_Click(sender As Object, e As EventArgs) Handles R13C1.Click

        CurrentRow = "13"
        CurrentColumn = "1"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R13C1.Text = CurrentNumber
            arrDone(192) = True
            arrValues(192) = CurrentNumber
        End If

    End Sub

    Private Sub R13C2_Click(sender As Object, e As EventArgs) Handles R13C2.Click

        CurrentRow = "13"
        CurrentColumn = "2"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R13C2.Text = CurrentNumber
            arrDone(193) = True
            arrValues(193) = CurrentNumber
        End If

    End Sub

    Private Sub R13C3_Click(sender As Object, e As EventArgs) Handles R13C3.Click

        CurrentRow = "13"
        CurrentColumn = "3"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R13C3.Text = CurrentNumber
            arrDone(194) = True
            arrValues(194) = CurrentNumber
        End If

    End Sub

    Private Sub R13C4_Click(sender As Object, e As EventArgs) Handles R13C4.Click

        CurrentRow = "13"
        CurrentColumn = "4"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R13C4.Text = CurrentNumber
            arrDone(195) = True
            arrValues(195) = CurrentNumber
        End If

    End Sub

    Private Sub R13C5_Click(sender As Object, e As EventArgs) Handles R13C5.Click

        CurrentRow = "13"
        CurrentColumn = "5"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R13C5.Text = CurrentNumber
            arrDone(196) = True
            arrValues(196) = CurrentNumber
        End If

    End Sub

    Private Sub R13C6_Click(sender As Object, e As EventArgs) Handles R13C6.Click

        CurrentRow = "13"
        CurrentColumn = "6"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R13C6.Text = CurrentNumber
            arrDone(197) = True
            arrValues(197) = CurrentNumber
        End If

    End Sub

    Private Sub R13C7_Click(sender As Object, e As EventArgs) Handles R13C7.Click

        CurrentRow = "13"
        CurrentColumn = "7"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R13C7.Text = CurrentNumber
            arrDone(198) = True
            arrValues(198) = CurrentNumber
        End If

    End Sub

    Private Sub R13C8_Click(sender As Object, e As EventArgs) Handles R13C8.Click

        CurrentRow = "13"
        CurrentColumn = "8"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R13C8.Text = CurrentNumber
            arrDone(199) = True
            arrValues(199) = CurrentNumber
        End If

    End Sub

    Private Sub R13C9_Click(sender As Object, e As EventArgs) Handles R13C9.Click

        CurrentRow = "13"
        CurrentColumn = "9"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R13C9.Text = CurrentNumber
            arrDone(200) = True
            arrValues(200) = CurrentNumber
        End If

    End Sub

    Private Sub R13C10_Click(sender As Object, e As EventArgs) Handles R13C10.Click

        CurrentRow = "13"
        CurrentColumn = "10"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R13C10.Text = CurrentNumber
            arrDone(201) = True
            arrValues(201) = CurrentNumber
        End If

    End Sub

    Private Sub R13C11_Click(sender As Object, e As EventArgs) Handles R13C11.Click

        CurrentRow = "13"
        CurrentColumn = "11"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R13C11.Text = CurrentNumber
            arrDone(202) = True
            arrValues(202) = CurrentNumber
        End If

    End Sub

    Private Sub R13C12_Click(sender As Object, e As EventArgs) Handles R13C12.Click

        CurrentRow = "13"
        CurrentColumn = "12"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R13C12.Text = CurrentNumber
            arrDone(203) = True
            arrValues(203) = CurrentNumber
        End If

    End Sub

    Private Sub R13C13_Click(sender As Object, e As EventArgs) Handles R13C13.Click

        CurrentRow = "13"
        CurrentColumn = "13"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R13C13.Text = CurrentNumber
            arrDone(204) = True
            arrValues(204) = CurrentNumber
        End If

    End Sub

    Private Sub R13C14_Click(sender As Object, e As EventArgs) Handles R13C14.Click

        CurrentRow = "13"
        CurrentColumn = "14"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R13C14.Text = CurrentNumber
            arrDone(205) = True
            arrValues(205) = CurrentNumber
        End If

    End Sub

    Private Sub R13C15_Click(sender As Object, e As EventArgs) Handles R13C15.Click

        CurrentRow = "13"
        CurrentColumn = "15"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R13C15.Text = CurrentNumber
            arrDone(206) = True
            arrValues(206) = CurrentNumber
        End If

    End Sub

    Private Sub R13C16_Click(sender As Object, e As EventArgs) Handles R13C16.Click

        CurrentRow = "13"
        CurrentColumn = "16"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R13C16.Text = CurrentNumber
            arrDone(207) = True
            arrValues(207) = CurrentNumber
        End If

    End Sub

    Private Sub R14C1_Click(sender As Object, e As EventArgs) Handles R14C1.Click

        CurrentRow = "14"
        CurrentColumn = "1"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R14C1.Text = CurrentNumber
            arrDone(208) = True
            arrValues(208) = CurrentNumber
        End If

    End Sub

    Private Sub R14C2_Click(sender As Object, e As EventArgs) Handles R14C2.Click

        CurrentRow = "14"
        CurrentColumn = "2"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R14C2.Text = CurrentNumber
            arrDone(209) = True
            arrValues(209) = CurrentNumber
        End If

    End Sub

    Private Sub R14C3_Click(sender As Object, e As EventArgs) Handles R14C3.Click

        CurrentRow = "14"
        CurrentColumn = "3"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R14C3.Text = CurrentNumber
            arrDone(210) = True
            arrValues(210) = CurrentNumber
        End If

    End Sub

    Private Sub R14C4_Click(sender As Object, e As EventArgs) Handles R14C4.Click

        CurrentRow = "14"
        CurrentColumn = "4"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R14C4.Text = CurrentNumber
            arrDone(211) = True
            arrValues(211) = CurrentNumber
        End If

    End Sub

    Private Sub R14C5_Click(sender As Object, e As EventArgs) Handles R14C5.Click

        CurrentRow = "14"
        CurrentColumn = "5"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R14C5.Text = CurrentNumber
            arrDone(212) = True
            arrValues(212) = CurrentNumber
        End If

    End Sub

    Private Sub R14C6_Click(sender As Object, e As EventArgs) Handles R14C6.Click

        CurrentRow = "14"
        CurrentColumn = "6"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R14C6.Text = CurrentNumber
            arrDone(213) = True
            arrValues(213) = CurrentNumber
        End If

    End Sub

    Private Sub R14C7_Click(sender As Object, e As EventArgs) Handles R14C7.Click

        CurrentRow = "14"
        CurrentColumn = "7"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R14C7.Text = CurrentNumber
            arrDone(214) = True
            arrValues(214) = CurrentNumber
        End If

    End Sub

    Private Sub R14C8_Click(sender As Object, e As EventArgs) Handles R14C8.Click

        CurrentRow = "14"
        CurrentColumn = "8"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R14C8.Text = CurrentNumber
            arrDone(215) = True
            arrValues(215) = CurrentNumber
        End If

    End Sub

    Private Sub R14C9_Click(sender As Object, e As EventArgs) Handles R14C9.Click

        CurrentRow = "14"
        CurrentColumn = "9"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R14C9.Text = CurrentNumber
            arrDone(216) = True
            arrValues(216) = CurrentNumber
        End If

    End Sub

    Private Sub R14C10_Click(sender As Object, e As EventArgs) Handles R14C10.Click

        CurrentRow = "14"
        CurrentColumn = "10"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R14C10.Text = CurrentNumber
            arrDone(217) = True
            arrValues(217) = CurrentNumber
        End If

    End Sub

    Private Sub R14C11_Click(sender As Object, e As EventArgs) Handles R14C11.Click

        CurrentRow = "14"
        CurrentColumn = "11"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R14C11.Text = CurrentNumber
            arrDone(218) = True
            arrValues(218) = CurrentNumber
        End If

    End Sub

    Private Sub R14C12_Click(sender As Object, e As EventArgs) Handles R14C12.Click

        CurrentRow = "14"
        CurrentColumn = "12"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R14C12.Text = CurrentNumber
            arrDone(219) = True
            arrValues(219) = CurrentNumber
        End If

    End Sub

    Private Sub R14C13_Click(sender As Object, e As EventArgs) Handles R14C13.Click

        CurrentRow = "14"
        CurrentColumn = "13"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R14C13.Text = CurrentNumber
            arrDone(220) = True
            arrValues(220) = CurrentNumber
        End If

    End Sub

    Private Sub R14C14_Click(sender As Object, e As EventArgs) Handles R14C14.Click

        CurrentRow = "14"
        CurrentColumn = "14"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R14C14.Text = CurrentNumber
            arrDone(221) = True
            arrValues(221) = CurrentNumber
        End If

    End Sub

    Private Sub R14C15_Click(sender As Object, e As EventArgs) Handles R14C15.Click

        CurrentRow = "14"
        CurrentColumn = "15"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R14C15.Text = CurrentNumber
            arrDone(222) = True
            arrValues(222) = CurrentNumber
        End If

    End Sub

    Private Sub R14C16_Click(sender As Object, e As EventArgs) Handles R14C16.Click

        CurrentRow = "14"
        CurrentColumn = "16"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R14C16.Text = CurrentNumber
            arrDone(223) = True
            arrValues(223) = CurrentNumber
        End If

    End Sub

    Private Sub R15C1_Click(sender As Object, e As EventArgs) Handles R15C1.Click

        CurrentRow = "15"
        CurrentColumn = "1"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R15C1.Text = CurrentNumber
            arrDone(224) = True
            arrValues(224) = CurrentNumber
        End If

    End Sub

    Private Sub R15C2_Click(sender As Object, e As EventArgs) Handles R15C2.Click

        CurrentRow = "15"
        CurrentColumn = "2"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R15C2.Text = CurrentNumber
            arrDone(225) = True
            arrValues(225) = CurrentNumber
        End If

    End Sub

    Private Sub R15C3_Click(sender As Object, e As EventArgs) Handles R15C3.Click

        CurrentRow = "15"
        CurrentColumn = "3"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R15C3.Text = CurrentNumber
            arrDone(226) = True
            arrValues(226) = CurrentNumber
        End If

    End Sub

    Private Sub R15C4_Click(sender As Object, e As EventArgs) Handles R15C4.Click

        CurrentRow = "15"
        CurrentColumn = "4"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R15C4.Text = CurrentNumber
            arrDone(227) = True
            arrValues(227) = CurrentNumber
        End If

    End Sub

    Private Sub R15C5_Click(sender As Object, e As EventArgs) Handles R15C5.Click

        CurrentRow = "15"
        CurrentColumn = "5"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R15C5.Text = CurrentNumber
            arrDone(228) = True
            arrValues(228) = CurrentNumber
        End If

    End Sub

    Private Sub R15C6_Click(sender As Object, e As EventArgs) Handles R15C6.Click

        CurrentRow = "15"
        CurrentColumn = "6"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R15C6.Text = CurrentNumber
            arrDone(229) = True
            arrValues(229) = CurrentNumber
        End If

    End Sub

    Private Sub R15C7_Click(sender As Object, e As EventArgs) Handles R15C7.Click

        CurrentRow = "15"
        CurrentColumn = "7"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R15C7.Text = CurrentNumber
            arrDone(230) = True
            arrValues(230) = CurrentNumber
        End If

    End Sub

    Private Sub R15C8_Click(sender As Object, e As EventArgs) Handles R15C8.Click

        CurrentRow = "15"
        CurrentColumn = "8"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R15C8.Text = CurrentNumber
            arrDone(231) = True
            arrValues(231) = CurrentNumber
        End If

    End Sub

    Private Sub R15C9_Click(sender As Object, e As EventArgs) Handles R15C9.Click

        CurrentRow = "15"
        CurrentColumn = "9"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R15C9.Text = CurrentNumber
            arrDone(232) = True
            arrValues(232) = CurrentNumber
        End If

    End Sub

    Private Sub R15C10_Click(sender As Object, e As EventArgs) Handles R15C10.Click

        CurrentRow = "15"
        CurrentColumn = "10"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R15C10.Text = CurrentNumber
            arrDone(233) = True
            arrValues(233) = CurrentNumber
        End If

    End Sub

    Private Sub R15C11_Click(sender As Object, e As EventArgs) Handles R15C11.Click

        CurrentRow = "15"
        CurrentColumn = "11"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R15C11.Text = CurrentNumber
            arrDone(234) = True
            arrValues(234) = CurrentNumber
        End If

    End Sub

    Private Sub R15C12_Click(sender As Object, e As EventArgs) Handles R15C12.Click

        CurrentRow = "15"
        CurrentColumn = "12"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R15C12.Text = CurrentNumber
            arrDone(235) = True
            arrValues(235) = CurrentNumber
        End If

    End Sub

    Private Sub R15C13_Click(sender As Object, e As EventArgs) Handles R15C13.Click

        CurrentRow = "15"
        CurrentColumn = "13"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R15C13.Text = CurrentNumber
            arrDone(236) = True
            arrValues(236) = CurrentNumber
        End If

    End Sub

    Private Sub R15C14_Click(sender As Object, e As EventArgs) Handles R15C14.Click

        CurrentRow = "15"
        CurrentColumn = "14"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R15C14.Text = CurrentNumber
            arrDone(237) = True
            arrValues(237) = CurrentNumber
        End If

    End Sub

    Private Sub R15C15_Click(sender As Object, e As EventArgs) Handles R15C15.Click

        CurrentRow = "15"
        CurrentColumn = "15"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R15C15.Text = CurrentNumber
            arrDone(238) = True
            arrValues(238) = CurrentNumber
        End If

    End Sub

    Private Sub R15C16_Click(sender As Object, e As EventArgs) Handles R15C16.Click

        CurrentRow = "15"
        CurrentColumn = "16"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R15C16.Text = CurrentNumber
            arrDone(239) = True
            arrValues(239) = CurrentNumber
        End If

    End Sub

    Private Sub R16C1_Click(sender As Object, e As EventArgs) Handles R16C1.Click

        CurrentRow = "16"
        CurrentColumn = "1"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R16C1.Text = CurrentNumber
            arrDone(240) = True
            arrValues(240) = CurrentNumber
        End If

    End Sub

    Private Sub R16C2_Click(sender As Object, e As EventArgs) Handles R16C2.Click

        CurrentRow = "16"
        CurrentColumn = "2"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R16C2.Text = CurrentNumber
            arrDone(241) = True
            arrValues(241) = CurrentNumber
        End If

    End Sub

    Private Sub R16C3_Click(sender As Object, e As EventArgs) Handles R16C3.Click

        CurrentRow = "16"
        CurrentColumn = "3"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R16C3.Text = CurrentNumber
            arrDone(242) = True
            arrValues(242) = CurrentNumber
        End If

    End Sub

    Private Sub R16C4_Click(sender As Object, e As EventArgs) Handles R16C4.Click

        CurrentRow = "16"
        CurrentColumn = "4"
        CurrentSquare = "13"

        If bRemoveANumberFromACell = False Then
            R16C4.Text = CurrentNumber
            arrDone(243) = True
            arrValues(243) = CurrentNumber
        End If

    End Sub

    Private Sub R16C5_Click(sender As Object, e As EventArgs) Handles R16C5.Click

        CurrentRow = "16"
        CurrentColumn = "5"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R16C5.Text = CurrentNumber
            arrDone(244) = True
            arrValues(244) = CurrentNumber
        End If

    End Sub

    Private Sub R16C6_Click(sender As Object, e As EventArgs) Handles R16C6.Click

        CurrentRow = "16"
        CurrentColumn = "6"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R16C6.Text = CurrentNumber
            arrDone(245) = True
            arrValues(245) = CurrentNumber
        End If

    End Sub

    Private Sub R16C7_Click(sender As Object, e As EventArgs) Handles R16C7.Click

        CurrentRow = "16"
        CurrentColumn = "7"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R16C7.Text = CurrentNumber
            arrDone(246) = True
            arrValues(246) = CurrentNumber
        End If

    End Sub

    Private Sub R16C8_Click(sender As Object, e As EventArgs) Handles R16C8.Click

        CurrentRow = "16"
        CurrentColumn = "8"
        CurrentSquare = "14"

        If bRemoveANumberFromACell = False Then
            R16C8.Text = CurrentNumber
            arrDone(247) = True
            arrValues(247) = CurrentNumber
        End If

    End Sub

    Private Sub R16C9_Click(sender As Object, e As EventArgs) Handles R16C9.Click

        CurrentRow = "16"
        CurrentColumn = "9"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R16C9.Text = CurrentNumber
            arrDone(248) = True
            arrValues(248) = CurrentNumber
        End If

    End Sub

    Private Sub R16C10_Click(sender As Object, e As EventArgs) Handles R16C10.Click

        CurrentRow = "16"
        CurrentColumn = "10"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R16C10.Text = CurrentNumber
            arrDone(249) = True
            arrValues(249) = CurrentNumber
        End If

    End Sub

    Private Sub R16C11_Click(sender As Object, e As EventArgs) Handles R16C11.Click

        CurrentRow = "16"
        CurrentColumn = "11"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R16C11.Text = CurrentNumber
            arrDone(250) = True
            arrValues(250) = CurrentNumber
        End If

    End Sub

    Private Sub R16C12_Click(sender As Object, e As EventArgs) Handles R16C12.Click

        CurrentRow = "16"
        CurrentColumn = "12"
        CurrentSquare = "15"

        If bRemoveANumberFromACell = False Then
            R16C12.Text = CurrentNumber
            arrDone(251) = True
            arrValues(251) = CurrentNumber
        End If

    End Sub

    Private Sub R16C13_Click(sender As Object, e As EventArgs) Handles R16C13.Click

        CurrentRow = "16"
        CurrentColumn = "13"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R16C13.Text = CurrentNumber
            arrDone(252) = True
            arrValues(252) = CurrentNumber
        End If

    End Sub

    Private Sub R16C14_Click(sender As Object, e As EventArgs) Handles R16C14.Click

        CurrentRow = "16"
        CurrentColumn = "14"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R16C14.Text = CurrentNumber
            arrDone(253) = True
            arrValues(253) = CurrentNumber
        End If

    End Sub

    Private Sub R16C15_Click(sender As Object, e As EventArgs) Handles R16C15.Click

        CurrentRow = "16"
        CurrentColumn = "15"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R16C15.Text = CurrentNumber
            arrDone(254) = True
            arrValues(254) = CurrentNumber
        End If

    End Sub

    Private Sub R16C16_Click(sender As Object, e As EventArgs) Handles R16C16.Click

        CurrentRow = "16"
        CurrentColumn = "16"
        CurrentSquare = "16"

        If bRemoveANumberFromACell = False Then
            R16C16.Text = CurrentNumber
            arrDone(255) = True
            arrValues(255) = CurrentNumber
        End If

    End Sub

    Private Sub ButtonUpdatePuzzle_Click_1(sender As Object, e As EventArgs) Handles ButtonUpdatePuzzle.Click
        UpdatePuzzle()

    End Sub

    Private Sub UpdatePuzzle()
        Me.TextBox1.Text = Me.TextBox1.Text & "Entering UpdatePuzzle " & vbCrLf

        'Go through all the arrays, if a cell is done, go to the next cell, otherwise
        'check to see if the row, column, or square are the same as the current row, column, or row
        'if they are the same, update that cell. 
        'At this point, don't check to see if there is only one number in the cell
        Dim i As Int16  'i is a generic counter
        'Dim l As Int16  'l is location of current number in the string
        Dim strTemp As String

        SaveCurrentSolution()
        'TextBox1.Text = TextBox1.Text & "In Sub ButtonUpdatePuzzle_Click" & vbCrLf
        'TextBox1.Text = TextBox1.Text & "CurrentNumber is " & CurrentNumber & vbCrLf
        If bRemoveANumberFromACell = True Then
            RemoveANumberFromACell()
        Else

            For i = 0 To 255

                If arrDone(i) = False Then
                    'Dont need the following because the arrValues was already set
                    'If arrReferences(i, 0) = CurrentRow And arrReferences(i, 1) = CurrentColumn And arrReferences(i, 2) = CurrentSquare Then
                    'arrValues(i) = CurrentNumber
                    'End If
                    'If i = 122 Then
                    'strTemp = arrValues(i)
                    'End If
                    If arrReferences(i, 0) = CurrentRow Or arrReferences(i, 1) = CurrentColumn Or arrReferences(i, 2) = CurrentSquare Then
                        strTemp = arrValues(i)
                        'TextBox1.Text = TextBox1.Text & "arrValues(i) is " & i & " " & strTemp & vbCrLf
                        'If any of the row, column or square references match, 
                        ' check to see if CurrentNumber is in the string, 
                        ' if CurrentNumber is in the string, remove the number from the current cell string
                        If InStr(strTemp, CurrentNumber) And Len(strTemp) > 1 Then
                            strTemp = Replace(strTemp, CurrentNumber, "")
                            arrValues(i) = strTemp
                            'ElseIf Len(strTemp) = 1 Then
                            '   arrValues(i) = strTemp
                            '  arrDone(i) = True
                        End If

                        'If Len(strTemp) = 1 Then
                        'arrValues(i) = strTemp
                        'arrDone(i) = True
                        'Else
                        '*** what was this for?
                        'End If

                        'The above works to identify all the right cells,
                        'and removes the current number from the cells!!!!


                    End If
                End If
            Next
        End If
        bRemoveANumberFromACell = False

        RefreshPuzzle()
    End Sub
    Private Sub RefreshPuzzle()
        Me.TextBox1.Text = Me.TextBox1.Text & "Entering RefreshPuzzle " & vbCrLf

        R1C1.Text = arrValues(0)
        R1C2.Text = arrValues(1)
        R1C3.Text = arrValues(2)
        R1C4.Text = arrValues(3)
        R1C5.Text = arrValues(4)
        R1C6.Text = arrValues(5)
        R1C7.Text = arrValues(6)
        R1C8.Text = arrValues(7)
        R1C9.Text = arrValues(8)
        R1C10.Text = arrValues(9)
        R1C11.Text = arrValues(10)
        R1C12.Text = arrValues(11)
        R1C13.Text = arrValues(12)
        R1C14.Text = arrValues(13)
        R1C15.Text = arrValues(14)
        R1C16.Text = arrValues(15)
        R2C1.Text = arrValues(16)
        R2C2.Text = arrValues(17)
        R2C3.Text = arrValues(18)
        R2C4.Text = arrValues(19)
        R2C5.Text = arrValues(20)
        R2C6.Text = arrValues(21)
        R2C7.Text = arrValues(22)
        R2C8.Text = arrValues(23)
        R2C9.Text = arrValues(24)
        R2C10.Text = arrValues(25)
        R2C11.Text = arrValues(26)
        R2C12.Text = arrValues(27)
        R2C13.Text = arrValues(28)
        R2C14.Text = arrValues(29)
        R2C15.Text = arrValues(30)
        R2C16.Text = arrValues(31)
        R3C1.Text = arrValues(32)
        R3C2.Text = arrValues(33)
        R3C3.Text = arrValues(34)
        R3C4.Text = arrValues(35)
        R3C5.Text = arrValues(36)
        R3C6.Text = arrValues(37)
        R3C7.Text = arrValues(38)
        R3C8.Text = arrValues(39)
        R3C9.Text = arrValues(40)
        R3C10.Text = arrValues(41)
        R3C11.Text = arrValues(42)
        R3C12.Text = arrValues(43)
        R3C13.Text = arrValues(44)
        R3C14.Text = arrValues(45)
        R3C15.Text = arrValues(46)
        R3C16.Text = arrValues(47)
        R4C1.Text = arrValues(48)
        R4C2.Text = arrValues(49)
        R4C3.Text = arrValues(50)
        R4C4.Text = arrValues(51)
        R4C5.Text = arrValues(52)
        R4C6.Text = arrValues(53)
        R4C7.Text = arrValues(54)
        R4C8.Text = arrValues(55)
        R4C9.Text = arrValues(56)
        R4C10.Text = arrValues(57)
        R4C11.Text = arrValues(58)
        R4C12.Text = arrValues(59)
        R4C13.Text = arrValues(60)
        R4C14.Text = arrValues(61)
        R4C15.Text = arrValues(62)
        R4C16.Text = arrValues(63)
        R5C1.Text = arrValues(64)
        R5C2.Text = arrValues(65)
        R5C3.Text = arrValues(66)
        R5C4.Text = arrValues(67)
        R5C5.Text = arrValues(68)
        R5C6.Text = arrValues(69)
        R5C7.Text = arrValues(70)
        R5C8.Text = arrValues(71)
        R5C9.Text = arrValues(72)
        R5C10.Text = arrValues(73)
        R5C11.Text = arrValues(74)
        R5C12.Text = arrValues(75)
        R5C13.Text = arrValues(76)
        R5C14.Text = arrValues(77)
        R5C15.Text = arrValues(78)
        R5C16.Text = arrValues(79)
        R6C1.Text = arrValues(80)
        R6C2.Text = arrValues(81)
        R6C3.Text = arrValues(82)
        R6C4.Text = arrValues(83)
        R6C5.Text = arrValues(84)
        R6C6.Text = arrValues(85)
        R6C7.Text = arrValues(86)
        R6C8.Text = arrValues(87)
        R6C9.Text = arrValues(88)
        R6C10.Text = arrValues(89)
        R6C11.Text = arrValues(90)
        R6C12.Text = arrValues(91)
        R6C13.Text = arrValues(92)
        R6C14.Text = arrValues(93)
        R6C15.Text = arrValues(94)
        R6C16.Text = arrValues(95)
        R7C1.Text = arrValues(96)
        R7C2.Text = arrValues(97)
        R7C3.Text = arrValues(98)
        R7C4.Text = arrValues(99)
        R7C5.Text = arrValues(100)
        R7C6.Text = arrValues(101)
        R7C7.Text = arrValues(102)
        R7C8.Text = arrValues(103)
        R7C9.Text = arrValues(104)
        R7C10.Text = arrValues(105)
        R7C11.Text = arrValues(106)
        R7C12.Text = arrValues(107)
        R7C13.Text = arrValues(108)
        R7C14.Text = arrValues(109)
        R7C15.Text = arrValues(110)
        R7C16.Text = arrValues(111)
        R8C1.Text = arrValues(112)
        R8C2.Text = arrValues(113)
        R8C3.Text = arrValues(114)
        R8C4.Text = arrValues(115)
        R8C5.Text = arrValues(116)
        R8C6.Text = arrValues(117)
        R8C7.Text = arrValues(118)
        R8C8.Text = arrValues(119)
        R8C9.Text = arrValues(120)
        R8C10.Text = arrValues(121)
        R8C11.Text = arrValues(122)
        R8C12.Text = arrValues(123)
        R8C13.Text = arrValues(124)
        R8C14.Text = arrValues(125)
        R8C15.Text = arrValues(126)
        R8C16.Text = arrValues(127)
        R9C1.Text = arrValues(128)
        R9C2.Text = arrValues(129)
        R9C3.Text = arrValues(130)
        R9C4.Text = arrValues(131)
        R9C5.Text = arrValues(132)
        R9C6.Text = arrValues(133)
        R9C7.Text = arrValues(134)
        R9C8.Text = arrValues(135)
        R9C9.Text = arrValues(136)
        R9C10.Text = arrValues(137)
        R9C11.Text = arrValues(138)
        R9C12.Text = arrValues(139)
        R9C13.Text = arrValues(140)
        R9C14.Text = arrValues(141)
        R9C15.Text = arrValues(142)
        R9C16.Text = arrValues(143)
        R10C1.Text = arrValues(144)
        R10C2.Text = arrValues(145)
        R10C3.Text = arrValues(146)
        R10C4.Text = arrValues(147)
        R10C5.Text = arrValues(148)
        R10C6.Text = arrValues(149)
        R10C7.Text = arrValues(150)
        R10C8.Text = arrValues(151)
        R10C9.Text = arrValues(152)
        R10C10.Text = arrValues(153)
        R10C11.Text = arrValues(154)
        R10C12.Text = arrValues(155)
        R10C13.Text = arrValues(156)
        R10C14.Text = arrValues(157)
        R10C15.Text = arrValues(158)
        R10C16.Text = arrValues(159)
        R11C1.Text = arrValues(160)
        R11C2.Text = arrValues(161)
        R11C3.Text = arrValues(162)
        R11C4.Text = arrValues(163)
        R11C5.Text = arrValues(164)
        R11C6.Text = arrValues(165)
        R11C7.Text = arrValues(166)
        R11C8.Text = arrValues(167)
        R11C9.Text = arrValues(168)
        R11C10.Text = arrValues(169)
        R11C11.Text = arrValues(170)
        R11C12.Text = arrValues(171)
        R11C13.Text = arrValues(172)
        R11C14.Text = arrValues(173)
        R11C15.Text = arrValues(174)
        R11C16.Text = arrValues(175)
        R12C1.Text = arrValues(176)
        R12C2.Text = arrValues(177)
        R12C3.Text = arrValues(178)
        R12C4.Text = arrValues(179)
        R12C5.Text = arrValues(180)
        R12C6.Text = arrValues(181)
        R12C7.Text = arrValues(182)
        R12C8.Text = arrValues(183)
        R12C9.Text = arrValues(184)
        R12C10.Text = arrValues(185)
        R12C11.Text = arrValues(186)
        R12C12.Text = arrValues(187)
        R12C13.Text = arrValues(188)
        R12C14.Text = arrValues(189)
        R12C15.Text = arrValues(190)
        R12C16.Text = arrValues(191)
        R13C1.Text = arrValues(192)
        R13C2.Text = arrValues(193)
        R13C3.Text = arrValues(194)
        R13C4.Text = arrValues(195)
        R13C5.Text = arrValues(196)
        R13C6.Text = arrValues(197)
        R13C7.Text = arrValues(198)
        R13C8.Text = arrValues(199)
        R13C9.Text = arrValues(200)
        R13C10.Text = arrValues(201)
        R13C11.Text = arrValues(202)
        R13C12.Text = arrValues(203)
        R13C13.Text = arrValues(204)
        R13C14.Text = arrValues(205)
        R13C15.Text = arrValues(206)
        R13C16.Text = arrValues(207)
        R14C1.Text = arrValues(208)
        R14C2.Text = arrValues(209)
        R14C3.Text = arrValues(210)
        R14C4.Text = arrValues(211)
        R14C5.Text = arrValues(212)
        R14C6.Text = arrValues(213)
        R14C7.Text = arrValues(214)
        R14C8.Text = arrValues(215)
        R14C9.Text = arrValues(216)
        R14C10.Text = arrValues(217)
        R14C11.Text = arrValues(218)
        R14C12.Text = arrValues(219)
        R14C13.Text = arrValues(220)
        R14C14.Text = arrValues(221)
        R14C15.Text = arrValues(222)
        R14C16.Text = arrValues(223)
        R15C1.Text = arrValues(224)
        R15C2.Text = arrValues(225)
        R15C3.Text = arrValues(226)
        R15C4.Text = arrValues(227)
        R15C5.Text = arrValues(228)
        R15C6.Text = arrValues(229)
        R15C7.Text = arrValues(230)
        R15C8.Text = arrValues(231)
        R15C9.Text = arrValues(232)
        R15C10.Text = arrValues(233)
        R15C11.Text = arrValues(234)
        R15C12.Text = arrValues(235)
        R15C13.Text = arrValues(236)
        R15C14.Text = arrValues(237)
        R15C15.Text = arrValues(238)
        R15C16.Text = arrValues(239)
        R16C1.Text = arrValues(240)
        R16C2.Text = arrValues(241)
        R16C3.Text = arrValues(242)
        R16C4.Text = arrValues(243)
        R16C5.Text = arrValues(244)
        R16C6.Text = arrValues(245)
        R16C7.Text = arrValues(246)
        R16C8.Text = arrValues(247)
        R16C9.Text = arrValues(248)
        R16C10.Text = arrValues(249)
        R16C11.Text = arrValues(250)
        R16C12.Text = arrValues(251)
        R16C13.Text = arrValues(252)
        R16C14.Text = arrValues(253)
        R16C15.Text = arrValues(254)
        R16C16.Text = arrValues(255)

        If arrDone(0) = True Then
            R1C1.ForeColor = Color.DarkRed
        ElseIf R1C1.ForeColor = Color.Black Then
        End If
        If arrDone(1) = True Then
            R1C2.ForeColor = Color.DarkRed
        ElseIf R1C2.ForeColor = Color.Black Then
        End If
        If arrDone(2) = True Then
            R1C3.ForeColor = Color.DarkRed
        ElseIf R1C3.ForeColor = Color.Black Then
        End If
        If arrDone(3) = True Then
            R1C4.ForeColor = Color.DarkRed
        ElseIf R1C4.ForeColor = Color.Black Then
        End If
        If arrDone(4) = True Then
            R1C5.ForeColor = Color.DarkRed
        ElseIf R1C5.ForeColor = Color.Black Then
        End If
        If arrDone(5) = True Then
            R1C6.ForeColor = Color.DarkRed
        ElseIf R1C6.ForeColor = Color.Black Then
        End If
        If arrDone(6) = True Then
            R1C7.ForeColor = Color.DarkRed
        ElseIf R1C7.ForeColor = Color.Black Then
        End If
        If arrDone(7) = True Then
            R1C8.ForeColor = Color.DarkRed
        ElseIf R1C8.ForeColor = Color.Black Then
        End If
        If arrDone(8) = True Then
            R1C9.ForeColor = Color.DarkRed
        ElseIf R1C9.ForeColor = Color.Black Then
        End If
        If arrDone(9) = True Then
            R1C10.ForeColor = Color.DarkRed
        ElseIf R1C10.ForeColor = Color.Black Then
        End If
        If arrDone(10) = True Then
            R1C11.ForeColor = Color.DarkRed
        ElseIf R1C11.ForeColor = Color.Black Then
        End If
        If arrDone(11) = True Then
            R1C12.ForeColor = Color.DarkRed
        ElseIf R1C11.ForeColor = Color.Black Then
        End If
        If arrDone(12) = True Then
            R1C13.ForeColor = Color.DarkRed
        ElseIf R1C13.ForeColor = Color.Black Then
        End If
        If arrDone(13) = True Then
            R1C14.ForeColor = Color.DarkRed
        ElseIf R1C14.ForeColor = Color.Black Then
        End If
        If arrDone(14) = True Then
            R1C15.ForeColor = Color.DarkRed
        ElseIf R1C15.ForeColor = Color.Black Then
        End If
        If arrDone(15) = True Then
            R1C16.ForeColor = Color.DarkRed
        ElseIf R1C16.ForeColor = Color.Black Then
        End If
        If arrDone(16) = True Then
            R2C1.ForeColor = Color.DarkRed
        ElseIf R2C1.ForeColor = Color.Black Then
        End If
        If arrDone(17) = True Then
            R2C2.ForeColor = Color.DarkRed
        ElseIf R2C2.ForeColor = Color.Black Then
        End If
        If arrDone(18) = True Then
            R2C3.ForeColor = Color.DarkRed
        ElseIf R2C3.ForeColor = Color.Black Then
        End If
        If arrDone(19) = True Then
            R2C4.ForeColor = Color.DarkRed
        ElseIf R2C4.ForeColor = Color.Black Then
        End If
        If arrDone(20) = True Then
            R2C5.ForeColor = Color.DarkRed
        ElseIf R2C5.ForeColor = Color.Black Then
        End If
        If arrDone(21) = True Then
            R2C6.ForeColor = Color.DarkRed
        ElseIf R2C6.ForeColor = Color.Black Then
        End If
        If arrDone(22) = True Then
            R2C7.ForeColor = Color.DarkRed
        ElseIf R2C7.ForeColor = Color.Black Then
        End If
        If arrDone(23) = True Then
            R2C8.ForeColor = Color.DarkRed
        ElseIf R2C8.ForeColor = Color.Black Then
        End If
        If arrDone(24) = True Then
            R2C9.ForeColor = Color.DarkRed
        ElseIf R2C9.ForeColor = Color.Black Then
        End If
        If arrDone(25) = True Then
            R2C10.ForeColor = Color.DarkRed
        ElseIf R2C10.ForeColor = Color.Black Then
        End If
        If arrDone(26) = True Then
            R2C11.ForeColor = Color.DarkRed
        ElseIf R2C11.ForeColor = Color.Black Then
        End If
        If arrDone(27) = True Then
            R2C12.ForeColor = Color.DarkRed
        ElseIf R2C11.ForeColor = Color.Black Then
        End If
        If arrDone(28) = True Then
            R2C13.ForeColor = Color.DarkRed
        ElseIf R2C13.ForeColor = Color.Black Then
        End If
        If arrDone(29) = True Then
            R2C14.ForeColor = Color.DarkRed
        ElseIf R2C14.ForeColor = Color.Black Then
        End If
        If arrDone(30) = True Then
            R2C15.ForeColor = Color.DarkRed
        ElseIf R2C15.ForeColor = Color.Black Then
        End If
        If arrDone(31) = True Then
            R2C16.ForeColor = Color.DarkRed
        ElseIf R2C16.ForeColor = Color.Black Then
        End If
        If arrDone(32) = True Then
            R3C1.ForeColor = Color.DarkRed
        ElseIf R3C1.ForeColor = Color.Black Then
        End If
        If arrDone(33) = True Then
            R3C2.ForeColor = Color.DarkRed
        ElseIf R3C2.ForeColor = Color.Black Then
        End If
        If arrDone(34) = True Then
            R3C3.ForeColor = Color.DarkRed
        ElseIf R3C3.ForeColor = Color.Black Then
        End If
        If arrDone(35) = True Then
            R3C4.ForeColor = Color.DarkRed
        ElseIf R3C4.ForeColor = Color.Black Then
        End If
        If arrDone(36) = True Then
            R3C5.ForeColor = Color.DarkRed
        ElseIf R3C5.ForeColor = Color.Black Then
        End If
        If arrDone(37) = True Then
            R3C6.ForeColor = Color.DarkRed
        ElseIf R3C6.ForeColor = Color.Black Then
        End If
        If arrDone(38) = True Then
            R3C7.ForeColor = Color.DarkRed
        ElseIf R3C7.ForeColor = Color.Black Then
        End If
        If arrDone(39) = True Then
            R3C8.ForeColor = Color.DarkRed
        ElseIf R3C8.ForeColor = Color.Black Then
        End If
        If arrDone(40) = True Then
            R3C9.ForeColor = Color.DarkRed
        ElseIf R3C9.ForeColor = Color.Black Then
        End If
        If arrDone(41) = True Then
            R3C10.ForeColor = Color.DarkRed
        ElseIf R3C10.ForeColor = Color.Black Then
        End If
        If arrDone(42) = True Then
            R3C11.ForeColor = Color.DarkRed
        ElseIf R3C11.ForeColor = Color.Black Then
        End If
        If arrDone(43) = True Then
            R3C12.ForeColor = Color.DarkRed
        ElseIf R3C11.ForeColor = Color.Black Then
        End If
        If arrDone(44) = True Then
            R3C13.ForeColor = Color.DarkRed
        ElseIf R3C13.ForeColor = Color.Black Then
        End If
        If arrDone(45) = True Then
            R3C14.ForeColor = Color.DarkRed
        ElseIf R3C14.ForeColor = Color.Black Then
        End If
        If arrDone(46) = True Then
            R3C15.ForeColor = Color.DarkRed
        ElseIf R3C15.ForeColor = Color.Black Then
        End If
        If arrDone(47) = True Then
            R3C16.ForeColor = Color.DarkRed
        ElseIf R3C16.ForeColor = Color.Black Then
        End If
        If arrDone(48) = True Then
            R4C1.ForeColor = Color.DarkRed
        ElseIf R4C1.ForeColor = Color.Black Then
        End If
        If arrDone(49) = True Then
            R4C2.ForeColor = Color.DarkRed
        ElseIf R4C2.ForeColor = Color.Black Then
        End If
        If arrDone(50) = True Then
            R4C3.ForeColor = Color.DarkRed
        ElseIf R4C3.ForeColor = Color.Black Then
        End If
        If arrDone(51) = True Then
            R4C4.ForeColor = Color.DarkRed
        ElseIf R4C4.ForeColor = Color.Black Then
        End If
        If arrDone(52) = True Then
            R4C5.ForeColor = Color.DarkRed
        ElseIf R4C5.ForeColor = Color.Black Then
        End If
        If arrDone(53) = True Then
            R4C6.ForeColor = Color.DarkRed
        ElseIf R4C6.ForeColor = Color.Black Then
        End If
        If arrDone(54) = True Then
            R4C7.ForeColor = Color.DarkRed
        ElseIf R4C7.ForeColor = Color.Black Then
        End If
        If arrDone(55) = True Then
            R4C8.ForeColor = Color.DarkRed
        ElseIf R4C8.ForeColor = Color.Black Then
        End If
        If arrDone(56) = True Then
            R4C9.ForeColor = Color.DarkRed
        ElseIf R4C9.ForeColor = Color.Black Then
        End If
        If arrDone(57) = True Then
            R4C10.ForeColor = Color.DarkRed
        ElseIf R4C10.ForeColor = Color.Black Then
        End If
        If arrDone(58) = True Then
            R4C11.ForeColor = Color.DarkRed
        ElseIf R4C11.ForeColor = Color.Black Then
        End If
        If arrDone(59) = True Then
            R4C12.ForeColor = Color.DarkRed
        ElseIf R4C11.ForeColor = Color.Black Then
        End If
        If arrDone(60) = True Then
            R4C13.ForeColor = Color.DarkRed
        ElseIf R4C13.ForeColor = Color.Black Then
        End If
        If arrDone(61) = True Then
            R4C14.ForeColor = Color.DarkRed
        ElseIf R4C14.ForeColor = Color.Black Then
        End If
        If arrDone(62) = True Then
            R4C15.ForeColor = Color.DarkRed
        ElseIf R4C15.ForeColor = Color.Black Then
        End If
        If arrDone(63) = True Then
            R4C16.ForeColor = Color.DarkRed
        ElseIf R4C16.ForeColor = Color.Black Then
        End If
        '***
        If arrDone(64) = True Then
            R5C1.ForeColor = Color.DarkRed
        ElseIf R5C1.ForeColor = Color.Black Then
        End If
        If arrDone(65) = True Then
            R5C2.ForeColor = Color.DarkRed
        ElseIf R5C2.ForeColor = Color.Black Then
        End If
        If arrDone(66) = True Then
            R5C3.ForeColor = Color.DarkRed
        ElseIf R5C3.ForeColor = Color.Black Then
        End If
        If arrDone(67) = True Then
            R5C4.ForeColor = Color.DarkRed
        ElseIf R5C4.ForeColor = Color.Black Then
        End If
        If arrDone(68) = True Then
            R5C5.ForeColor = Color.DarkRed
        ElseIf R5C5.ForeColor = Color.Black Then
        End If
        If arrDone(69) = True Then
            R5C6.ForeColor = Color.DarkRed
        ElseIf R5C6.ForeColor = Color.Black Then
        End If
        If arrDone(70) = True Then
            R5C7.ForeColor = Color.DarkRed
        ElseIf R5C7.ForeColor = Color.Black Then
        End If
        If arrDone(71) = True Then
            R5C8.ForeColor = Color.DarkRed
        ElseIf R5C8.ForeColor = Color.Black Then
        End If
        If arrDone(72) = True Then
            R5C9.ForeColor = Color.DarkRed
        ElseIf R5C9.ForeColor = Color.Black Then
        End If
        If arrDone(73) = True Then
            R5C10.ForeColor = Color.DarkRed
        ElseIf R5C10.ForeColor = Color.Black Then
        End If
        If arrDone(74) = True Then
            R5C11.ForeColor = Color.DarkRed
        ElseIf R5C11.ForeColor = Color.Black Then
        End If
        If arrDone(75) = True Then
            R5C12.ForeColor = Color.DarkRed
        ElseIf R5C11.ForeColor = Color.Black Then
        End If
        If arrDone(76) = True Then
            R5C13.ForeColor = Color.DarkRed
        ElseIf R5C13.ForeColor = Color.Black Then
        End If
        If arrDone(77) = True Then
            R5C14.ForeColor = Color.DarkRed
        ElseIf R5C14.ForeColor = Color.Black Then
        End If
        If arrDone(78) = True Then
            R5C15.ForeColor = Color.DarkRed
        ElseIf R5C15.ForeColor = Color.Black Then
        End If
        If arrDone(79) = True Then
            R5C16.ForeColor = Color.DarkRed
        ElseIf R5C16.ForeColor = Color.Black Then
        End If
        If arrDone(80) = True Then
            R6C1.ForeColor = Color.DarkRed
        ElseIf R6C1.ForeColor = Color.Black Then
        End If
        If arrDone(81) = True Then
            R6C2.ForeColor = Color.DarkRed
        ElseIf R6C2.ForeColor = Color.Black Then
        End If
        If arrDone(82) = True Then
            R6C3.ForeColor = Color.DarkRed
        ElseIf R6C3.ForeColor = Color.Black Then
        End If
        If arrDone(83) = True Then
            R6C4.ForeColor = Color.DarkRed
        ElseIf R6C4.ForeColor = Color.Black Then
        End If
        If arrDone(84) = True Then
            R6C5.ForeColor = Color.DarkRed
        ElseIf R6C5.ForeColor = Color.Black Then
        End If
        If arrDone(85) = True Then
            R6C6.ForeColor = Color.DarkRed
        ElseIf R6C6.ForeColor = Color.Black Then
        End If
        If arrDone(86) = True Then
            R6C7.ForeColor = Color.DarkRed
        ElseIf R6C7.ForeColor = Color.Black Then
        End If
        If arrDone(87) = True Then
            R6C8.ForeColor = Color.DarkRed
        ElseIf R6C8.ForeColor = Color.Black Then
        End If
        If arrDone(88) = True Then
            R6C9.ForeColor = Color.DarkRed
        ElseIf R6C9.ForeColor = Color.Black Then
        End If
        If arrDone(89) = True Then
            R6C10.ForeColor = Color.DarkRed
        ElseIf R6C10.ForeColor = Color.Black Then
        End If
        If arrDone(90) = True Then
            R6C11.ForeColor = Color.DarkRed
        ElseIf R6C11.ForeColor = Color.Black Then
        End If
        If arrDone(91) = True Then
            R6C12.ForeColor = Color.DarkRed
        ElseIf R6C11.ForeColor = Color.Black Then
        End If
        If arrDone(92) = True Then
            R6C13.ForeColor = Color.DarkRed
        ElseIf R6C13.ForeColor = Color.Black Then
        End If
        If arrDone(93) = True Then
            R6C14.ForeColor = Color.DarkRed
        ElseIf R6C14.ForeColor = Color.Black Then
        End If
        If arrDone(94) = True Then
            R6C15.ForeColor = Color.DarkRed
        ElseIf R6C15.ForeColor = Color.Black Then
        End If
        If arrDone(95) = True Then
            R6C16.ForeColor = Color.DarkRed
        ElseIf R6C16.ForeColor = Color.Black Then
        End If
        If arrDone(96) = True Then
            R7C1.ForeColor = Color.DarkRed
        ElseIf R7C1.ForeColor = Color.Black Then
        End If
        If arrDone(97) = True Then
            R7C2.ForeColor = Color.DarkRed
        ElseIf R7C2.ForeColor = Color.Black Then
        End If
        If arrDone(98) = True Then
            R7C3.ForeColor = Color.DarkRed
        ElseIf R7C3.ForeColor = Color.Black Then
        End If
        If arrDone(99) = True Then
            R7C4.ForeColor = Color.DarkRed
        ElseIf R7C4.ForeColor = Color.Black Then
        End If
        If arrDone(100) = True Then
            R7C5.ForeColor = Color.DarkRed
        ElseIf R7C5.ForeColor = Color.Black Then
        End If
        If arrDone(101) = True Then
            R7C6.ForeColor = Color.DarkRed
        ElseIf R7C6.ForeColor = Color.Black Then
        End If
        If arrDone(102) = True Then
            R7C7.ForeColor = Color.DarkRed
        ElseIf R7C7.ForeColor = Color.Black Then
        End If
        If arrDone(103) = True Then
            R7C8.ForeColor = Color.DarkRed
        ElseIf R7C8.ForeColor = Color.Black Then
        End If
        If arrDone(104) = True Then
            R7C9.ForeColor = Color.DarkRed
        ElseIf R7C9.ForeColor = Color.Black Then
        End If
        If arrDone(105) = True Then
            R7C10.ForeColor = Color.DarkRed
        ElseIf R7C10.ForeColor = Color.Black Then
        End If
        If arrDone(106) = True Then
            R7C11.ForeColor = Color.DarkRed
        ElseIf R7C11.ForeColor = Color.Black Then
        End If
        If arrDone(107) = True Then
            R7C12.ForeColor = Color.DarkRed
        ElseIf R7C12.ForeColor = Color.Black Then
        End If
        If arrDone(108) = True Then
            R7C13.ForeColor = Color.DarkRed
        ElseIf R7C13.ForeColor = Color.Black Then
        End If
        If arrDone(109) = True Then
            R7C14.ForeColor = Color.DarkRed
        ElseIf R7C14.ForeColor = Color.Black Then
        End If
        If arrDone(110) = True Then
            R7C15.ForeColor = Color.DarkRed
        ElseIf R7C15.ForeColor = Color.Black Then
        End If
        If arrDone(111) = True Then
            R7C16.ForeColor = Color.DarkRed
        ElseIf R7C16.ForeColor = Color.Black Then
        End If
        If arrDone(112) = True Then
            R8C1.ForeColor = Color.DarkRed
        ElseIf R8C1.ForeColor = Color.Black Then
        End If
        If arrDone(113) = True Then
            R8C2.ForeColor = Color.DarkRed
        ElseIf R8C2.ForeColor = Color.Black Then
        End If
        If arrDone(114) = True Then
            R8C3.ForeColor = Color.DarkRed
        ElseIf R8C3.ForeColor = Color.Black Then
        End If
        If arrDone(115) = True Then
            R8C4.ForeColor = Color.DarkRed
        ElseIf R8C4.ForeColor = Color.Black Then
        End If
        If arrDone(116) = True Then
            R8C5.ForeColor = Color.DarkRed
        ElseIf R8C5.ForeColor = Color.Black Then
        End If
        If arrDone(117) = True Then
            R8C6.ForeColor = Color.DarkRed
        ElseIf R8C6.ForeColor = Color.Black Then
        End If
        If arrDone(118) = True Then
            R8C7.ForeColor = Color.DarkRed
        ElseIf R8C7.ForeColor = Color.Black Then
        End If
        If arrDone(119) = True Then
            R8C8.ForeColor = Color.DarkRed
        ElseIf R8C8.ForeColor = Color.Black Then
        End If
        If arrDone(120) = True Then
            R8C9.ForeColor = Color.DarkRed
        ElseIf R8C9.ForeColor = Color.Black Then
        End If
        If arrDone(121) = True Then
            R8C10.ForeColor = Color.DarkRed
        ElseIf R8C10.ForeColor = Color.Black Then
        End If
        If arrDone(122) = True Then
            R8C11.ForeColor = Color.DarkRed
        ElseIf R8C11.ForeColor = Color.Black Then
        End If
        If arrDone(123) = True Then
            R8C12.ForeColor = Color.DarkRed
        ElseIf R8C12.ForeColor = Color.Black Then
        End If
        If arrDone(124) = True Then
            R8C13.ForeColor = Color.DarkRed
        ElseIf R8C13.ForeColor = Color.Black Then
        End If
        If arrDone(125) = True Then
            R8C14.ForeColor = Color.DarkRed
        ElseIf R8C14.ForeColor = Color.Black Then
        End If
        If arrDone(126) = True Then
            R8C15.ForeColor = Color.DarkRed
        ElseIf R8C15.ForeColor = Color.Black Then
        End If
        If arrDone(127) = True Then
            R8C16.ForeColor = Color.DarkRed
        ElseIf R8C16.ForeColor = Color.Black Then
        End If
        If arrDone(128) = True Then
            R9C1.ForeColor = Color.DarkRed
        ElseIf R9C1.ForeColor = Color.Black Then
        End If
        If arrDone(129) = True Then
            R9C2.ForeColor = Color.DarkRed
        ElseIf R9C2.ForeColor = Color.Black Then
        End If
        If arrDone(130) = True Then
            R9C3.ForeColor = Color.DarkRed
        ElseIf R9C3.ForeColor = Color.Black Then
        End If
        If arrDone(131) = True Then
            R9C4.ForeColor = Color.DarkRed
        ElseIf R9C4.ForeColor = Color.Black Then
        End If
        If arrDone(132) = True Then
            R9C5.ForeColor = Color.DarkRed
        ElseIf R9C5.ForeColor = Color.Black Then
        End If
        If arrDone(133) = True Then
            R9C6.ForeColor = Color.DarkRed
        ElseIf R9C6.ForeColor = Color.Black Then
        End If
        If arrDone(134) = True Then
            R9C7.ForeColor = Color.DarkRed
        ElseIf R9C7.ForeColor = Color.Black Then
        End If
        If arrDone(135) = True Then
            R9C8.ForeColor = Color.DarkRed
        ElseIf R9C8.ForeColor = Color.Black Then
        End If
        If arrDone(136) = True Then
            R9C9.ForeColor = Color.DarkRed
        ElseIf R9C9.ForeColor = Color.Black Then
        End If
        If arrDone(137) = True Then
            R9C10.ForeColor = Color.DarkRed
        ElseIf R9C10.ForeColor = Color.Black Then
        End If
        If arrDone(138) = True Then
            R9C11.ForeColor = Color.DarkRed
        ElseIf R9C11.ForeColor = Color.Black Then
        End If
        If arrDone(139) = True Then
            R9C12.ForeColor = Color.DarkRed
        ElseIf R9C12.ForeColor = Color.Black Then
        End If
        If arrDone(140) = True Then
            R9C13.ForeColor = Color.DarkRed
        ElseIf R9C13.ForeColor = Color.Black Then
        End If
        If arrDone(141) = True Then
            R9C14.ForeColor = Color.DarkRed
        ElseIf R9C14.ForeColor = Color.Black Then
        End If
        If arrDone(142) = True Then
            R9C15.ForeColor = Color.DarkRed
        ElseIf R9C15.ForeColor = Color.Black Then
        End If
        If arrDone(143) = True Then
            R9C16.ForeColor = Color.DarkRed
        ElseIf R9C16.ForeColor = Color.Black Then
        End If
        If arrDone(144) = True Then
            R10C1.ForeColor = Color.DarkRed
        ElseIf R10C1.ForeColor = Color.Black Then
        End If
        If arrDone(145) = True Then
            R10C2.ForeColor = Color.DarkRed
        ElseIf R10C2.ForeColor = Color.Black Then
        End If
        If arrDone(146) = True Then
            R10C3.ForeColor = Color.DarkRed
        ElseIf R10C3.ForeColor = Color.Black Then
        End If
        If arrDone(147) = True Then
            R10C4.ForeColor = Color.DarkRed
        ElseIf R10C4.ForeColor = Color.Black Then
        End If
        If arrDone(148) = True Then
            R10C5.ForeColor = Color.DarkRed
        ElseIf R10C5.ForeColor = Color.Black Then
        End If
        If arrDone(149) = True Then
            R10C6.ForeColor = Color.DarkRed
        ElseIf R10C6.ForeColor = Color.Black Then
        End If
        If arrDone(150) = True Then
            R10C7.ForeColor = Color.DarkRed
        ElseIf R10C7.ForeColor = Color.Black Then
        End If
        If arrDone(151) = True Then
            R10C8.ForeColor = Color.DarkRed
        ElseIf R10C8.ForeColor = Color.Black Then
        End If
        If arrDone(152) = True Then
            R10C9.ForeColor = Color.DarkRed
        ElseIf R10C9.ForeColor = Color.Black Then
        End If
        If arrDone(153) = True Then
            R10C10.ForeColor = Color.DarkRed
        ElseIf R10C10.ForeColor = Color.Black Then
        End If
        If arrDone(154) = True Then
            R10C11.ForeColor = Color.DarkRed
        ElseIf R10C11.ForeColor = Color.Black Then
        End If
        If arrDone(155) = True Then
            R10C12.ForeColor = Color.DarkRed
        ElseIf R10C12.ForeColor = Color.Black Then
        End If
        If arrDone(156) = True Then
            R10C13.ForeColor = Color.DarkRed
        ElseIf R10C13.ForeColor = Color.Black Then
        End If
        If arrDone(157) = True Then
            R10C14.ForeColor = Color.DarkRed
        ElseIf R10C14.ForeColor = Color.Black Then
        End If
        If arrDone(158) = True Then
            R10C15.ForeColor = Color.DarkRed
        ElseIf R10C15.ForeColor = Color.Black Then
        End If
        If arrDone(159) = True Then
            R10C16.ForeColor = Color.DarkRed
        ElseIf R10C16.ForeColor = Color.Black Then
        End If
        If arrDone(160) = True Then
            R11C1.ForeColor = Color.DarkRed
        ElseIf R11C1.ForeColor = Color.Black Then
        End If
        If arrDone(161) = True Then
            R11C2.ForeColor = Color.DarkRed
        ElseIf R11C2.ForeColor = Color.Black Then
        End If
        If arrDone(162) = True Then
            R11C3.ForeColor = Color.DarkRed
        ElseIf R11C3.ForeColor = Color.Black Then
        End If
        If arrDone(163) = True Then
            R11C4.ForeColor = Color.DarkRed
        ElseIf R11C4.ForeColor = Color.Black Then
        End If
        If arrDone(164) = True Then
            R11C5.ForeColor = Color.DarkRed
        ElseIf R11C5.ForeColor = Color.Black Then
        End If
        If arrDone(165) = True Then
            R11C6.ForeColor = Color.DarkRed
        ElseIf R11C6.ForeColor = Color.Black Then
        End If
        If arrDone(166) = True Then
            R11C7.ForeColor = Color.DarkRed
        ElseIf R11C7.ForeColor = Color.Black Then
        End If
        If arrDone(167) = True Then
            R11C8.ForeColor = Color.DarkRed
        ElseIf R11C8.ForeColor = Color.Black Then
        End If
        If arrDone(168) = True Then
            R11C9.ForeColor = Color.DarkRed
        ElseIf R11C9.ForeColor = Color.Black Then
        End If
        If arrDone(169) = True Then
            R11C10.ForeColor = Color.DarkRed
        ElseIf R11C10.ForeColor = Color.Black Then
        End If
        If arrDone(170) = True Then
            R11C11.ForeColor = Color.DarkRed
        ElseIf R11C11.ForeColor = Color.Black Then
        End If
        If arrDone(171) = True Then
            R11C12.ForeColor = Color.DarkRed
        ElseIf R11C12.ForeColor = Color.Black Then
        End If
        If arrDone(172) = True Then
            R11C13.ForeColor = Color.DarkRed
        ElseIf R11C13.ForeColor = Color.Black Then
        End If
        If arrDone(173) = True Then
            R11C14.ForeColor = Color.DarkRed
        ElseIf R11C14.ForeColor = Color.Black Then
        End If
        If arrDone(174) = True Then
            R11C15.ForeColor = Color.DarkRed
        ElseIf R11C15.ForeColor = Color.Black Then
        End If
        If arrDone(175) = True Then
            R11C16.ForeColor = Color.DarkRed
        ElseIf R11C16.ForeColor = Color.Black Then
        End If
        If arrDone(176) = True Then
            R12C1.ForeColor = Color.DarkRed
        ElseIf R12C1.ForeColor = Color.Black Then
        End If
        If arrDone(177) = True Then
            R12C2.ForeColor = Color.DarkRed
        ElseIf R12C2.ForeColor = Color.Black Then
        End If
        If arrDone(178) = True Then
            R12C3.ForeColor = Color.DarkRed
        ElseIf R12C3.ForeColor = Color.Black Then
        End If
        If arrDone(179) = True Then
            R12C4.ForeColor = Color.DarkRed
        ElseIf R12C4.ForeColor = Color.Black Then
        End If
        If arrDone(180) = True Then
            R12C5.ForeColor = Color.DarkRed
        ElseIf R12C5.ForeColor = Color.Black Then
        End If
        If arrDone(181) = True Then
            R12C6.ForeColor = Color.DarkRed
        ElseIf R12C6.ForeColor = Color.Black Then
        End If
        If arrDone(182) = True Then
            R12C7.ForeColor = Color.DarkRed
        ElseIf R12C7.ForeColor = Color.Black Then
        End If
        If arrDone(183) = True Then
            R12C8.ForeColor = Color.DarkRed
        ElseIf R12C8.ForeColor = Color.Black Then
        End If
        If arrDone(184) = True Then
            R12C9.ForeColor = Color.DarkRed
        ElseIf R12C9.ForeColor = Color.Black Then
        End If
        If arrDone(185) = True Then
            R12C10.ForeColor = Color.DarkRed
        ElseIf R12C10.ForeColor = Color.Black Then
        End If
        If arrDone(186) = True Then
            R12C11.ForeColor = Color.DarkRed
        ElseIf R12C11.ForeColor = Color.Black Then
        End If
        If arrDone(187) = True Then
            R12C12.ForeColor = Color.DarkRed
        ElseIf R12C12.ForeColor = Color.Black Then
        End If
        If arrDone(188) = True Then
            R12C13.ForeColor = Color.DarkRed
        ElseIf R12C13.ForeColor = Color.Black Then
        End If
        If arrDone(189) = True Then
            R12C14.ForeColor = Color.DarkRed
        ElseIf R12C14.ForeColor = Color.Black Then
        End If
        If arrDone(190) = True Then
            R12C15.ForeColor = Color.DarkRed
        ElseIf R12C15.ForeColor = Color.Black Then
        End If
        If arrDone(191) = True Then
            R12C16.ForeColor = Color.DarkRed
        ElseIf R12C16.ForeColor = Color.Black Then
        End If
        If arrDone(192) = True Then
            R13C1.ForeColor = Color.DarkRed
        ElseIf R13C1.ForeColor = Color.Black Then
        End If
        If arrDone(193) = True Then
            R13C2.ForeColor = Color.DarkRed
        ElseIf R13C2.ForeColor = Color.Black Then
        End If
        If arrDone(194) = True Then
            R13C3.ForeColor = Color.DarkRed
        ElseIf R13C3.ForeColor = Color.Black Then
        End If
        If arrDone(195) = True Then
            R13C4.ForeColor = Color.DarkRed
        ElseIf R13C4.ForeColor = Color.Black Then
        End If
        If arrDone(196) = True Then
            R13C5.ForeColor = Color.DarkRed
        ElseIf R13C5.ForeColor = Color.Black Then
        End If
        If arrDone(197) = True Then
            R13C6.ForeColor = Color.DarkRed
        ElseIf R13C6.ForeColor = Color.Black Then
        End If
        If arrDone(198) = True Then
            R13C7.ForeColor = Color.DarkRed
        ElseIf R13C7.ForeColor = Color.Black Then
        End If
        If arrDone(199) = True Then
            R13C8.ForeColor = Color.DarkRed
        ElseIf R13C8.ForeColor = Color.Black Then
        End If
        If arrDone(200) = True Then
            R13C9.ForeColor = Color.DarkRed
        ElseIf R13C9.ForeColor = Color.Black Then
        End If
        If arrDone(201) = True Then
            R13C10.ForeColor = Color.DarkRed
        ElseIf R13C10.ForeColor = Color.Black Then
        End If
        If arrDone(202) = True Then
            R13C11.ForeColor = Color.DarkRed
        ElseIf R13C11.ForeColor = Color.Black Then
        End If
        If arrDone(203) = True Then
            R13C12.ForeColor = Color.DarkRed
        ElseIf R13C12.ForeColor = Color.Black Then
        End If
        If arrDone(204) = True Then
            R13C13.ForeColor = Color.DarkRed
        ElseIf R13C13.ForeColor = Color.Black Then
        End If
        If arrDone(205) = True Then
            R13C14.ForeColor = Color.DarkRed
        ElseIf R13C14.ForeColor = Color.Black Then
        End If
        If arrDone(206) = True Then
            R13C15.ForeColor = Color.DarkRed
        ElseIf R13C15.ForeColor = Color.Black Then
        End If
        If arrDone(207) = True Then
            R13C16.ForeColor = Color.DarkRed
        ElseIf R13C16.ForeColor = Color.Black Then
        End If
        If arrDone(208) = True Then
            R14C1.ForeColor = Color.DarkRed
        ElseIf R14C1.ForeColor = Color.Black Then
        End If
        If arrDone(209) = True Then
            R14C2.ForeColor = Color.DarkRed
        ElseIf R14C2.ForeColor = Color.Black Then
        End If
        If arrDone(210) = True Then
            R14C3.ForeColor = Color.DarkRed
        ElseIf R14C3.ForeColor = Color.Black Then
        End If
        If arrDone(211) = True Then
            R14C4.ForeColor = Color.DarkRed
        ElseIf R14C4.ForeColor = Color.Black Then
        End If
        If arrDone(212) = True Then
            R14C5.ForeColor = Color.DarkRed
        ElseIf R14C5.ForeColor = Color.Black Then
        End If
        If arrDone(213) = True Then
            R14C6.ForeColor = Color.DarkRed
        ElseIf R14C6.ForeColor = Color.Black Then
        End If
        If arrDone(214) = True Then
            R14C7.ForeColor = Color.DarkRed
        ElseIf R14C7.ForeColor = Color.Black Then
        End If
        If arrDone(215) = True Then
            R14C8.ForeColor = Color.DarkRed
        ElseIf R14C8.ForeColor = Color.Black Then
        End If
        If arrDone(216) = True Then
            R14C9.ForeColor = Color.DarkRed
        ElseIf R14C9.ForeColor = Color.Black Then
        End If
        If arrDone(217) = True Then
            R14C10.ForeColor = Color.DarkRed
        ElseIf R14C10.ForeColor = Color.Black Then
        End If
        If arrDone(218) = True Then
            R14C11.ForeColor = Color.DarkRed
        ElseIf R14C11.ForeColor = Color.Black Then
        End If
        If arrDone(219) = True Then
            R14C12.ForeColor = Color.DarkRed
        ElseIf R14C12.ForeColor = Color.Black Then
        End If
        If arrDone(220) = True Then
            R14C13.ForeColor = Color.DarkRed
        ElseIf R14C13.ForeColor = Color.Black Then
        End If
        If arrDone(221) = True Then
            R14C14.ForeColor = Color.DarkRed
        ElseIf R14C14.ForeColor = Color.Black Then
        End If
        If arrDone(222) = True Then
            R14C15.ForeColor = Color.DarkRed
        ElseIf R14C15.ForeColor = Color.Black Then
        End If
        If arrDone(223) = True Then
            R14C16.ForeColor = Color.DarkRed
        ElseIf R14C16.ForeColor = Color.Black Then
        End If
        If arrDone(224) = True Then
            R15C1.ForeColor = Color.DarkRed
        ElseIf R15C1.ForeColor = Color.Black Then
        End If
        If arrDone(225) = True Then
            R15C2.ForeColor = Color.DarkRed
        ElseIf R15C2.ForeColor = Color.Black Then
        End If
        If arrDone(226) = True Then
            R15C3.ForeColor = Color.DarkRed
        ElseIf R15C3.ForeColor = Color.Black Then
        End If
        If arrDone(227) = True Then
            R15C4.ForeColor = Color.DarkRed
        ElseIf R15C4.ForeColor = Color.Black Then
        End If
        If arrDone(228) = True Then
            R15C5.ForeColor = Color.DarkRed
        ElseIf R15C5.ForeColor = Color.Black Then
        End If
        If arrDone(229) = True Then
            R15C6.ForeColor = Color.DarkRed
        ElseIf R15C6.ForeColor = Color.Black Then
        End If
        If arrDone(230) = True Then
            R15C7.ForeColor = Color.DarkRed
        ElseIf R15C7.ForeColor = Color.Black Then
        End If
        If arrDone(231) = True Then
            R15C8.ForeColor = Color.DarkRed
        ElseIf R15C8.ForeColor = Color.Black Then
        End If
        If arrDone(232) = True Then
            R15C9.ForeColor = Color.DarkRed
        ElseIf R15C9.ForeColor = Color.Black Then
        End If
        If arrDone(233) = True Then
            R15C10.ForeColor = Color.DarkRed
        ElseIf R15C10.ForeColor = Color.Black Then
        End If
        If arrDone(234) = True Then
            R15C11.ForeColor = Color.DarkRed
        ElseIf R15C11.ForeColor = Color.Black Then
        End If
        If arrDone(235) = True Then
            R15C12.ForeColor = Color.DarkRed
        ElseIf R15C12.ForeColor = Color.Black Then
        End If
        If arrDone(236) = True Then
            R15C13.ForeColor = Color.DarkRed
        ElseIf R15C13.ForeColor = Color.Black Then
        End If
        If arrDone(237) = True Then
            R15C14.ForeColor = Color.DarkRed
        ElseIf R15C14.ForeColor = Color.Black Then
        End If
        If arrDone(238) = True Then
            R15C15.ForeColor = Color.DarkRed
        ElseIf R15C15.ForeColor = Color.Black Then
        End If
        If arrDone(239) = True Then
            R15C16.ForeColor = Color.DarkRed
        ElseIf R15C16.ForeColor = Color.Black Then
        End If
        If arrDone(240) = True Then
            R16C1.ForeColor = Color.DarkRed
        ElseIf R16C1.ForeColor = Color.Black Then
        End If
        If arrDone(241) = True Then
            R16C2.ForeColor = Color.DarkRed
        ElseIf R16C2.ForeColor = Color.Black Then
        End If
        If arrDone(242) = True Then
            R16C3.ForeColor = Color.DarkRed
        ElseIf R16C3.ForeColor = Color.Black Then
        End If
        If arrDone(243) = True Then
            R16C4.ForeColor = Color.DarkRed
        ElseIf R16C4.ForeColor = Color.Black Then
        End If
        If arrDone(244) = True Then
            R16C5.ForeColor = Color.DarkRed
        ElseIf R16C5.ForeColor = Color.Black Then
        End If
        If arrDone(245) = True Then
            R16C6.ForeColor = Color.DarkRed
        ElseIf R16C6.ForeColor = Color.Black Then
        End If
        If arrDone(246) = True Then
            R16C7.ForeColor = Color.DarkRed
        ElseIf R16C7.ForeColor = Color.Black Then
        End If
        If arrDone(247) = True Then
            R16C8.ForeColor = Color.DarkRed
        ElseIf R16C8.ForeColor = Color.Black Then
        End If
        If arrDone(248) = True Then
            R16C9.ForeColor = Color.DarkRed
        ElseIf R16C9.ForeColor = Color.Black Then
        End If
        If arrDone(249) = True Then
            R16C10.ForeColor = Color.DarkRed
        ElseIf R16C10.ForeColor = Color.Black Then
        End If
        If arrDone(250) = True Then
            R16C11.ForeColor = Color.DarkRed
        ElseIf R16C11.ForeColor = Color.Black Then
        End If
        If arrDone(251) = True Then
            R16C12.ForeColor = Color.DarkRed
        ElseIf R16C12.ForeColor = Color.Black Then
        End If
        If arrDone(252) = True Then
            R16C13.ForeColor = Color.DarkRed
        ElseIf R16C13.ForeColor = Color.Black Then
        End If
        If arrDone(253) = True Then
            R16C14.ForeColor = Color.DarkRed
        ElseIf R16C14.ForeColor = Color.Black Then
        End If
        If arrDone(254) = True Then
            R16C15.ForeColor = Color.DarkRed
        ElseIf R16C15.ForeColor = Color.Black Then
        End If
        If arrDone(255) = True Then
            R16C16.ForeColor = Color.DarkRed
        ElseIf R16C16.ForeColor = Color.Black Then
        End If

        PuzzleStatus()

    End Sub

    Private Sub ButtonSaveSolution1_Click(sender As Object, e As EventArgs) Handles ButtonSaveSolution1.Click
        SaveSolution1()
    End Sub

    Private Sub SaveSolution1()
        Dim i As Int16

        For i = 0 To 255
            arrDone1(i) = arrDone(i)
            arrValues1(i) = arrValues(i)
            'arrDone(i) = False
            'arrValues(i) = StartingString
            'TextBox1.Text = TextBox1.Text & i & ";" & arrDone(i) & ";" & arrValues(i) & ";"
            TextBox1.Text = TextBox1.Text & "arrDone(" & i & ") = " & Chr(34) & arrDone(i) & Chr(34) & vbCrLf
            TextBox1.Text = TextBox1.Text & "arrValues(" & i & ") = " & Chr(34) & arrValues(i) & Chr(34) & vbCrLf
        Next i
        TextBox1.Text = TextBox1.Text & Chr(34) & TextBox3.Text & Chr(34) & vbCrLf


    End Sub

    Private Sub ButtonRestoreSolution1_Click(sender As Object, e As EventArgs) Handles ButtonRestoreSolution1.Click
        RestoreSolution1()
    End Sub
    Private Sub RestoreSolution1()
        Dim i As Int16

        For i = 0 To 255
            arrDone(i) = arrDone1(i)
            arrValues(i) = arrValues1(i)
        Next i
        RefreshPuzzle()
        Me.TextBox3.Text = Hints

    End Sub

    ' ****** The following Sub is not called from anywhere ******
    ' ****** It would process the text from the text box ******
    Private Sub RestoreSolution()
        Dim strSolution As String
        Dim i As Int16
        Dim strTemp As String
        'Dim l As Int16
        Dim iLoc As Int16
        Dim intCount As Int16
        Dim strDone As String
        Dim strValue As String


        strSolution = TextBox1.Text
        TextBox1.Text = ""
        intCount = Len(strSolution) + 1
        strTemp = strSolution
        Do Until strTemp = ""
            iLoc = InStr(strTemp, ";")
            'TextBox1.Text = ""
            'TextBox1.Text = "This Function Is Not yet operational."
            'TextBox1.Text = TextBox1.Text & strSolution
            If iLoc = 0 Then iLoc = 1
            End

            i = Mid(strTemp, 1, iLoc - 1)
            'Trim off the used part
            strTemp = Mid(strTemp, iLoc + 1)
            iLoc = InStr(strTemp, ";")
            strDone = Mid(strTemp, 1, iLoc - 1)
            strTemp = Mid(strTemp, iLoc + 1)
            iLoc = InStr(strTemp, ";")
            strValue = Mid(strTemp, 1, iLoc - 1)
            arrValues(i) = strValue
            iLoc = InStr(strTemp, ";")
            strTemp = Mid(strTemp, iLoc + 1)
            TextBox1.Text = TextBox1.Text & i & "+" & strDone & "+" & arrValues(i) & vbCrLf
        Loop
        'l = Len(strTemp)
        RefreshPuzzle()

    End Sub

    Private Sub ButtonSaveSolution2_Click(sender As Object, e As EventArgs) Handles ButtonSaveSolution2.Click
        Dim i As Int16

        For i = 0 To 255
            arrDone2(i) = arrDone(i)
            arrValues2(i) = arrValues(i)
            TextBox1.Text = TextBox1.Text & "arrDone(" & i & ") = " & Chr(34) & arrDone(i) & Chr(34) & vbCrLf
            TextBox1.Text = TextBox1.Text & "arrValues(" & i & ") = " & Chr(34) & arrValues(i) & Chr(34) & vbCrLf
        Next i
    End Sub

    Private Sub ButtonRestoreSolution2_Click(sender As Object, e As EventArgs) Handles ButtonRestoreSolution2.Click
        Dim i As Int16

        For i = 0 To 255
            arrDone(i) = arrDone2(i)
            arrValues(i) = arrValues2(i)
        Next i
        RefreshPuzzle()
    End Sub
    Private Sub ButtonRemoveANumber_Click_1(sender As Object, e As EventArgs) Handles ButtonRemoveANumber.Click
        bRemoveANumberFromACell = True
        'RemoveANumberFromACell()
    End Sub

    Public Sub RemoveANumberFromACell()
        Me.TextBox1.Text = Me.TextBox1.Text & "Entering RemoveANumberFromACell " & vbCrLf

        Dim i As Int16
        Dim strTemp As String

        bRemoveANumberFromACell = True

        For i = 0 To 255
            If arrDone(i) = False Then
                If arrReferences(i, 0) = CurrentRow And arrReferences(i, 1) = CurrentColumn And arrReferences(i, 2) = CurrentSquare Then
                    strTemp = arrValues(i)
                    'TextBox1.Text = TextBox1.Text & "arrValues(i) Is " & i & " " & strTemp & vbCrLf
                    'If any of the row, column or square references match, 
                    ' check to see if CurrentNumber is in the string, 
                    ' if CurrentNumber is in the string, remove the number from the current cell string
                    If InStr(strTemp, CurrentNumber) Then
                        strTemp = Replace(strTemp, CurrentNumber, "")
                        arrValues(i) = strTemp
                        'The above works to identify all the right cells,
                        'and removes the current number from the cells!!!!

                    End If
                End If
            End If
        Next
        bRemoveANumberFromACell = False
    End Sub

    'Private Sub CheckBoxAutoComplete_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxAutoComplete.CheckedChanged
    'bAutoComplete = True
    'If Me.CheckBoxAutoComplete.Checked = True Then
    '       bAutoComplete = True
    'Me.TextBox1.Text = Me.TextBox1.Text & "bAutoComplete  Is " & bAutoComplete
    'ElseIf Me.CheckBoxAutoComplete.Checked = False Then
    '        bAutoComplete = False
    'Me.TextBox1.Text = Me.TextBox1.Text & "bAutoComplete  Is " & bAutoComplete
    'End If
    'End Sub

    Private Sub CheckForOnlyOneNumber()
        Me.TextBox1.Text = Me.TextBox1.Text & "Entering CheckForOnlyOneNumber " & vbCrLf
        Dim i As Int16
        'Dim j As Int16
        Dim strTemp As String

        'If there is only one number in a cell, mark it done, update its value, and refresh the puzzle
        'Also, if a cell causes an update, run through the puzzle again looking for another cell. Thus, reset the counter "j"
        'CheckForOnlyOneNumber()
        'For j = 0 To 1
        For i = 0 To 255
            If arrDone(i) = False Then
                strTemp = arrValues(i)
                If Len(strTemp) = 1 Then
                    If i = 74 Then
                        Me.TextBox1.Text = TextBox1.Text & "Only number is " & CurrentRow & ", " & CurrentColumn & ", " & CurrentNumber & vbCrLf
                    End If
                    CurrentRow = arrReferences(i, 0)
                    CurrentColumn = arrReferences(i, 1)
                    CurrentSquare = arrReferences(i, 2)
                    CurrentNumber = strTemp
                    Me.TextBox1.Text = TextBox1.Text & "Only number is " & CurrentRow & ", " & CurrentColumn & ", " & CurrentNumber & vbCrLf

                    arrDone(i) = True
                    arrValues(i) = CurrentNumber
                    UpdatePuzzle()
                    RefreshPuzzle()
                    Exit Sub
                End If
            End If
        Next
    End Sub

    Private Sub CheckForOnlyNumberInARow()
        Me.TextBox1.Text = Me.TextBox1.Text & "Entering CheckForOnlyNumberInARow " & vbCrLf

        Dim hChr As String     'The current number being checked
        Dim i As Int16     'Cell iteration counter
        Dim j As Int16     'Row counter
        'Dim k As Int16     'Number counter
        Dim r As Int16     '
        Dim m As Int16      'Length of string in the row, col, or sqare
        Dim n As Int16
        Dim p As Int16
        Dim z As Int16

        Dim strTemp As String
        Dim strStart As String
        Dim strTempStr As String

        strTemp = ""
        strStart = "0123456789ABCDEF"
        strTempStr = ""

        'Me.TextBox1.Text = TextBox1.Text & "Entering CheckForOnlyNumberInARow" & vbCrLf
        'While the program is being checked, display text messages in the textbox to show what is happening.
        'This sub starts the process by building a an empty string for each row and then a list of all strings in each row

        'Check each character from 1 through 9 for regular Sudoku or from 0 through F for Monster Sudoku
        'For each number "h" from 1 through 9, 
        ' Start by processing all cells -- I will need to do this several times, once for rows, columns and squares in other subs

        'arrStringsRemaining(2, 16)

        'need to ensure assStringsRemaining strings are empty--otherwise they keep being added to prior results
        ' The 0 in arrStringsRemaining(0, j) refers to rows, 1 in arrStringsRemaining(1, j) refers to columns 
        'And 2 in arrStringsRemaining(2, j) refers to squares
        'Although the array is "0" based, I started with 1 to make it easier to remember
        For j = 1 To 16
            arrStringsRemaining(0, j) = ""  'Ensure all row strings are reset to empty strings
        Next
        'Me.TextBox1.Text = TextBox1.Text & "Rows strings have been set to null" & vbCrLf
        For i = 0 To 255                    'Cells arrays are also 0 based and I used 0 base for the iteration
            For j = 1 To 16      'Start by doing the first row and go through row 16
                'Me.TextBox1.Text = TextBox1.Text & "arrReferences(" & i & ", 0) is " & arrReferences(i, 0) & vbCrLf
                If arrReferences(i, 0) = j Then 'If this cell is in the current row, process it
                    If arrDone(i) = False Then  'If the cell is not done, add its string of numbers to the arrStringsRemaining string
                        strTemp = arrStringsRemaining(0, j) 'Get the current string values from the arrStringsRemaining string
                        arrStringsRemaining(0, j) = strTemp & arrValues(i)  'Add the new values to the old ones
                    End If
                End If
                'After I get the rows to work, I will do a similar process for columns and squares
                'arrReferences(0, 1) = "1"
                'arrReferences(0, 2) = "1"
            Next

        Next
        'At this point, every row should have a list of all the numbers from all the cells in that row
        'For j = 1 To 16
        'Me.TextBox1.Text = TextBox1.Text & arrStringsRemaining(0, j) & vbCrLf
        'Next
        j = 1   'Start processing the 'now filled' rows

        p = Len(strStart)   'Get the length of the starting string so I can iterate over it's length

        'Now that the total strings have been made for each row, I need to find unique numbers, e.g, count < 2
        'Iterate through the starting string

        For r = 1 To Len(strStart)
            hChr = Mid(strStart, r, 1)
            'k = 0

            For j = 0 To 15      'to 16
                'k = 0
                ' need to check to see if the cell is done
                'For i = 0 To 255
                'If arrReferences(i, 0) = j And arrDone(i) = False Then
                strTempStr = arrStringsRemaining(0, j)
                m = Len(strTempStr)
                If InStr(strTempStr, hChr) Then
                    n = InStr(strTempStr, hChr)
                    'k = 1
                    strTempStr = Mid(strTempStr, n + 1)
                    'now check to see if there is a 2nd occurance of that character
                    If InStr(strTempStr, hChr) Then     'check to see if the character is in the rest of the string
                        'k = k + 1
                    Else    'There is only one instance of the character in the row,so update the cell with that number
                        'Me.TextBox1.Text = """"
                        z = j + 1   'Need to adjust row number for zero base
                        Me.TextBox1.Text = TextBox1.Text & "There is only one instance of character " & hChr & " in row " & z & vbCrLf
                        Me.TextBox2.Text = TextBox2.Text & "Row " & j & " " & hChr & vbCrLf

                        CurrentNumber = hChr
                        CurrentRow = j
                        UpdateCharInRow()
                        'RefreshPuzzle()
                        Exit Sub
                        'Me.TextBox1.Text = TextBox1.Text & arrStringsRemaining(0, j) & vbCrLf
                    End If
                End If
                'Report the character as the only one in the row
            Next
        Next


        '
        'Check each row for 1 through 9
        'start with "j" row = 1 
        'Check every cell from "i" = 0 through 80 -- actually check the arrDone(i) and the arrValues(i)
        'Iterate through all the cells "i" checking for the current row number "j"
        'and the occurance of a cell with that number that is done arrDone(i), or more than one occurance of the number in that row.
        'If there is a cell in the row that is done, go to the next row number and reset the number counter to 0
        'If there is an occurance of that number in a cell, increment the number counter by 1, and set "l" to "i"
        'so I can mark that cell as having the current number in it.
        'If the number counter gets To 2, go to the next row number and reset the number counter "k" to 0.
        'If the end of a row is reached and the number counter is 1, then there is only one occurance of that number in that row
        'and it is not in a cell that is "done"
        'Therefore, set CurrentNumber = h   ;  the number being checked,
        'set arrDone(i) to done,
        'set arrValue(i) to CurrentNumber
        'reset number counter "k" to 0
        'While the program is being checked, display text messages in the textbox to show what is happening.
        'Since I can only determine if there is only one occurance when the end of the row is reached, 
        'the row counter can be incremented.



    End Sub

    Private Sub CheckForOnlyNumberInAColumn()
        Me.TextBox1.Text = Me.TextBox1.Text & "Entering CheckForOnlyNumberInAColumn " & vbCrLf

        Dim hChr As String     'The current number being checked
        Dim i As Int16     'Cell iteration counter
        Dim j As Int16     'Row counter
        Dim k As Int16     'Number counter
        Dim r As Int16     '
        Dim m As Int16      'Length of string in the row, col, or sqare
        Dim n As Int16
        Dim p As Int16
        Dim strTemp As String
        Dim strStart As String
        Dim strTempStr As String

        strTemp = ""
        strStart = "0123456789ABCDEF"
        strTempStr = ""

        'While the program is being checked, display text messages in the textbox to show what is happening.
        'This sub starts the process by building a list of all strings in the row

        'Check each character from 1 through 9 or from 0 through F
        'For each number "h" from 1 through 9, 
        ' Start by processing all cells -- I will need to do this several times, once for rows, columns and squares

        'arrStringsRemaining(2, 15)

        'need to ensure assStringsRemaining strings are empty--otherwise they keep being added to prior results
        For j = 1 To 16
            arrStringsRemaining(1, j) = ""
        Next

        For i = 0 To 255
            For j = 1 To 16      'Start by doing one the first row
                If arrReferences(i, 1) = j Then 'If this cell is in the current row, process it
                    If arrDone(i) = False Then
                        strTemp = arrStringsRemaining(1, j)             'Set strTemp to value current column string 
                        arrStringsRemaining(1, j) = strTemp & arrValues(i)  'add the current string to the prior string
                    End If
                End If
                'After I get the rows to work, I will do a similar process for columns and squares
                'arrReferences(0, 1) = "1"
                'arrReferences(0, 2) = "1"
            Next

        Next

        'For j = 0 To 15
        'Me.TextBox1.Text = TextBox1.Text & arrStringsRemaining(1, j) & vbCrLf
        'Next
        j = 1

        p = Len(strStart)

        'Now that the total strings have been made for each row, I need to find unique numbers, e.g, count < 2
        'Iterate through the starting string

        For r = 1 To Len(strStart)
            hChr = Mid(strStart, r, 1)
            k = 0

            For j = 0 To 16      'to 16
                k = 0
                ' need to check to see if the cell is done
                'For i = 0 To 255
                'If arrReferences(i, 0) = j And arrDone(i) = False Then
                strTempStr = arrStringsRemaining(1, j)
                m = Len(strTempStr)
                If InStr(strTempStr, hChr) Then
                    n = InStr(strTempStr, hChr)
                    k = 1
                    strTempStr = Mid(strTempStr, n + 1)
                    'now check to see if there is a 2nd occurance of that character
                    If InStr(strTempStr, hChr) Then     'check to see if the character is in the rest of the string
                        k = k + 1                       'there are at least 2 instances of that char in the string
                    Else    'There is only one instance of that character in the string
                        Me.TextBox1.Text = TextBox1.Text & "There is only one instance of character " & hChr & " in column " & j & vbCrLf
                        Me.TextBox2.Text = TextBox2.Text & "Column " & j & " " & hChr & vbCrLf

                        CurrentNumber = hChr
                        CurrentColumn = j
                        UpdateCharInColumn()
                        Exit Sub
                    End If
                End If
                'Report the character as the only one in the row
            Next
        Next

    End Sub


    Private Sub CheckForOnlyNumberInASquare()
        Me.TextBox1.Text = Me.TextBox1.Text & "Entering CheckForOnlyNumberInASquare " & vbCrLf

        Dim hChr As String     'The current number being checked
        Dim i As Int16     'Cell iteration counter
        Dim j As Int16     'Row counter
        Dim k As Int16     'Number counter
        Dim r As Int16     '
        Dim m As Int16      'Length of string in the row, col, or sqare
        Dim n As Int16
        Dim p As Int16
        Dim strTemp As String
        Dim strStart As String
        Dim strTempStr As String

        strTemp = ""
        strStart = "0123456789ABCDEF"
        strTempStr = ""

        'While the program is being checked, display text messages in the textbox to show what is happening.
        'This sub starts the process by building a list of all strings in the row

        'Check each character from 1 through 9 or from 0 through F
        'For each number "h" from 1 through 9, 
        ' Start by processing all cells -- I will need to do this several times, once for rows, columns and squares

        'arrStringsRemaining(2, 15)

        'need to ensure assStringsRemaining strings are empty--otherwise they keep being added to prior results
        For j = 1 To 16
            arrStringsRemaining(2, j) = ""
        Next

        For i = 0 To 255
            For j = 1 To 16      'Start by doing one the first row
                If arrReferences(i, 2) = j Then 'If this cell is in the current row, process it
                    If arrDone(i) = False Then
                        strTemp = arrStringsRemaining(2, j)
                        arrStringsRemaining(2, j) = strTemp & arrValues(i)
                    End If
                End If
                'After I get the rows to work, I will do a similar process for columns and squares
                'arrReferences(0, 1) = "1"
                'arrReferences(0, 2) = "1"
            Next

        Next

        'For j = 1 To 16
        'Me.TextBox1.Text = TextBox1.Text & arrStringsRemaining(2, j) & vbCrLf
        'Next
        j = 1

        p = Len(strStart)

        'Now that the total strings have been made for each row, I need to find unique numbers, e.g, count < 2
        'Iterate through the starting string

        For r = 1 To Len(strStart)
            hChr = Mid(strStart, r, 1)
            k = 0

            For j = 1 To 16      'to 16
                k = 0
                ' need to check to see if the cell is done
                'For i = 0 To 255
                'If arrReferences(i, 0) = j And arrDone(i) = False Then
                strTempStr = arrStringsRemaining(2, j)
                m = Len(strTempStr)
                If InStr(strTempStr, hChr) Then
                    n = InStr(strTempStr, hChr)
                    k = 1
                    strTempStr = Mid(strTempStr, n + 1)
                    'now check to see if there is a 2nd occurance of that character
                    If InStr(strTempStr, hChr) Then     'check to see if the character is in the rest of the string
                        k = k + 1
                    Else    'do nothing because that character is already done for that square; but print the following
                        Me.TextBox1.Text = TextBox1.Text & "There is only one instance of character " & hChr & " in square " & j & vbCrLf
                        Me.TextBox2.Text = TextBox2.Text & "Square " & j & " " & hChr & vbCrLf
                        CurrentNumber = hChr
                        CurrentSquare = j
                        UpdateCharInSquare()
                        Exit Sub
                    End If
                End If
            Next
        Next

    End Sub
    'Private Sub ButtonClearTextBox1_Click(sender As Object, e As EventArgs) Handles ButtonClearTextbox1.Click
    '    TextBox1.Text = ""
    'End Sub




    'Private Sub ButtonAutoCmplete_Click(sender As Object, e As EventArgs) Handles ButtonAutoCmplete.Click
    '    CheckForOnlyOneNumber() ' This sub checks every cell for not done and string length = 1
    '    CheckForOnlyNumberInARow()
    '    CheckForOnlyNumberInAColumn()
    'CheckForOnlyNumberInASquare()

    'End Sub
    Private Sub ButtonCheckFor1Number_Click(sender As Object, e As EventArgs) Handles ButtonCheckFor1Number.Click
        CheckForOnlyOneNumber()
        'CheckForOnlyNumberInARow()
        'CheckForOnlyNumberInAColumn()
        'CheckForOnlyNumberInASquare()
    End Sub

    Private Sub PuzzleStatus()
        Dim i As Int16
        'Dim j As Int16
        'Dim k As Int16

        D = 0
        R = 0

        For i = 0 To 255
            If arrDone(i) = True Then
                D = D + 1
            Else
                R = R + 1    '256 - 1
            End If
        Next
        'If j + k <> 256 Then
        'MsgBox "error in number of cells counted"
        'End If
        Me.TextCellsDone.Text = D
        Me.TextCellsRemaining.Text = R
    End Sub
    Private Sub UpdateCharInRow()
        Me.TextBox1.Text = Me.TextBox1.Text & "Entering UpdateCharInRow " & vbCrLf

        Dim i As Int16
        'Dim j As Int16
        Dim strTemp As String


        For i = 0 To 255
            If arrDone(i) = False Then
                'For j = 1 To 16      'Start by doing one the first row   arrReferences(0, 0)
                If arrReferences(i, 0) = CurrentRow Then 'If this cell is in the current row, process it
                    strTemp = arrValues(i)     'Get the string from the cell
                    If InStr(strTemp, CurrentNumber) Then
                        CurrentRow = arrReferences(i, 0)
                        CurrentColumn = arrReferences(i, 1)
                        CurrentSquare = arrReferences(i, 2)
                        arrDone(i) = True
                        arrValues(i) = CurrentNumber
                        UpdatePuzzle()
                        RefreshPuzzle()
                        Exit Sub
                    End If
                End If
            End If
            'After I get the rows to work, I will do a similar process for columns and squares
            'arrReferences(0, 1) = "1"
            'arrReferences(0, 2) = "1"
        Next

    End Sub

    Private Sub UpdateCharInColumn()
        'Me.TextBox1.Text = Me.TextBox1.Text & "Entering UpdateCharInColumn " & vbCrLf

        Dim i As Int16
        'Dim j As Int16
        Dim strTemp As String


        For i = 0 To 255
            If arrDone(i) = False Then
                'For j = 1 To 16      'Start by doing one the first row   arrReferences(0, 0)
                If arrReferences(i, 1) = CurrentColumn Then 'If this cell is in the current column, process it
                    strTemp = arrValues(i)     'Get the string from the cell
                    If InStr(strTemp, CurrentNumber) Then
                        CurrentRow = arrReferences(i, 0)
                        CurrentColumn = arrReferences(i, 1)
                        CurrentSquare = arrReferences(i, 2)
                        arrValues(i) = CurrentNumber
                        arrDone(i) = True
                        UpdatePuzzle()
                        RefreshPuzzle()
                        Exit Sub
                    End If
                End If
            End If
            'After I get the rows to work, I will do a similar process for columns and squares
            'arrReferences(0, 1) = "1"
            'arrReferences(0, 2) = "1"
        Next

    End Sub

    Private Sub UpdateCharInSquare()
        'Me.TextBox1.Text = Me.TextBox1.Text & "Entering UpdateCharInSquare " & vbCrLf

        Dim i As Int16
        ' Dim j As Int16
        Dim strTemp As String


        For i = 0 To 255
            If arrDone(i) = False Then
                'For j = 1 To 16      'Start by doing one the first row   arrReferences(0, 0)
                If arrReferences(i, 2) = CurrentSquare Then 'If this cell is in the current row, process it
                    strTemp = arrValues(i)     'Get the string from the cell
                    If InStr(strTemp, CurrentNumber) Then
                        CurrentRow = arrReferences(i, 0)
                        CurrentColumn = arrReferences(i, 1)
                        CurrentSquare = arrReferences(i, 2)
                        arrValues(i) = CurrentNumber
                        arrDone(i) = True
                        UpdatePuzzle()
                        RefreshPuzzle()
                        Exit Sub
                    End If
                End If
            End If
            'After I get the rows to work, I will do a similar process for columns and squares
            'arrReferences(0, 1) = "1"
            'arrReferences(0, 2) = "1"
        Next

    End Sub

    Private Sub ButtonCheckRows_Click(sender As Object, e As EventArgs) Handles ButtonCheckRows.Click
        CheckForOnlyNumberInARow()
    End Sub
    Private Sub ButtonCheckColumns_Click(sender As Object, e As EventArgs) Handles ButtonCheckColumns.Click

        CheckForOnlyNumberInAColumn()
    End Sub

    Private Sub ButtonCheckSquare_Click(sender As Object, e As EventArgs) Handles ButtonCheckSquare.Click

        CheckForOnlyNumberInASquare()
    End Sub

    Private Sub ButtonLoadSolution1_Click_1(sender As Object, e As EventArgs) Handles ButtonLoadSolution1.Click
        arrDone(0) = "True"
        arrValues(0) = "2"
        arrDone(1) = "False"
        arrValues(1) = "0568A"
        arrDone(2) = "False"
        arrValues(2) = "6789AB"
        arrDone(3) = "False"
        arrValues(3) = "056AB"
        arrDone(4) = "False"
        arrValues(4) = "689B"
        arrDone(5) = "True"
        arrValues(5) = "3"
        arrDone(6) = "True"
        arrValues(6) = "D"
        arrDone(7) = "False"
        arrValues(7) = "468AB"
        arrDone(8) = "True"
        arrValues(8) = "1"
        arrDone(9) = "False"
        arrValues(9) = "5789"
        arrDone(10) = "True"
        arrValues(10) = "C"
        arrDone(11) = "False"
        arrValues(11) = "0578"
        arrDone(12) = "False"
        arrValues(12) = "0459A"
        arrDone(13) = "False"
        arrValues(13) = "0459B"
        arrDone(14) = "True"
        arrValues(14) = "F"
        arrDone(15) = "True"
        arrValues(15) = "E"
        arrDone(16) = "False"
        arrValues(16) = "13589BD"
        arrDone(17) = "False"
        arrValues(17) = "0358A"
        arrDone(18) = "False"
        arrValues(18) = "189ABD"
        arrDone(19) = "False"
        arrValues(19) = "0135ABD"
        arrDone(20) = "False"
        arrValues(20) = "1289BC"
        arrDone(21) = "True"
        arrValues(21) = "7"
        arrDone(22) = "False"
        arrValues(22) = "1289C"
        arrDone(23) = "False"
        arrValues(23) = "18ABC"
        arrDone(24) = "False"
        arrValues(24) = "389A"
        arrDone(25) = "True"
        arrValues(25) = "F"
        arrDone(26) = "True"
        arrValues(26) = "4"
        arrDone(27) = "True"
        arrValues(27) = "E"
        arrDone(28) = "True"
        arrValues(28) = "6"
        arrDone(29) = "False"
        arrValues(29) = "0359BC"
        arrDone(30) = "False"
        arrValues(30) = "09BC"
        arrDone(31) = "False"
        arrValues(31) = "039AC"
        arrDone(32) = "True"
        arrValues(32) = "E"
        arrDone(33) = "False"
        arrValues(33) = "0368AF"
        arrDone(34) = "True"
        arrValues(34) = "C"
        arrDone(35) = "False"
        arrValues(35) = "036A"
        arrDone(36) = "False"
        arrValues(36) = "689F"
        arrDone(37) = "False"
        arrValues(37) = "469AF"
        arrDone(38) = "False"
        arrValues(38) = "4689F"
        arrDone(39) = "True"
        arrValues(39) = "5"
        arrDone(40) = "True"
        arrValues(40) = "2"
        arrDone(41) = "False"
        arrValues(41) = "389"
        arrDone(42) = "False"
        arrValues(42) = "03A"
        arrDone(43) = "True"
        arrValues(43) = "B"
        arrDone(44) = "True"
        arrValues(44) = "D"
        arrDone(45) = "False"
        arrValues(45) = "0349"
        arrDone(46) = "True"
        arrValues(46) = "1"
        arrDone(47) = "True"
        arrValues(47) = "7"
        arrDone(48) = "False"
        arrValues(48) = "13579BDF"
        arrDone(49) = "True"
        arrValues(49) = "4"
        arrDone(50) = "False"
        arrValues(50) = "179ABDF"
        arrDone(51) = "False"
        arrValues(51) = "135ABD"
        arrDone(52) = "False"
        arrValues(52) = "19BCF"
        arrDone(53) = "True"
        arrValues(53) = "0"
        arrDone(54) = "True"
        arrValues(54) = "E"
        arrDone(55) = "False"
        arrValues(55) = "1ABC"
        arrDone(56) = "False"
        arrValues(56) = "379A"
        arrDone(57) = "False"
        arrValues(57) = "3579D"
        arrDone(58) = "True"
        arrValues(58) = "6"
        arrDone(59) = "False"
        arrValues(59) = "357D"
        arrDone(60) = "True"
        arrValues(60) = "2"
        arrDone(61) = "False"
        arrValues(61) = "359BC"
        arrDone(62) = "True"
        arrValues(62) = "8"
        arrDone(63) = "False"
        arrValues(63) = "39AC"
        arrDone(64) = "False"
        arrValues(64) = "89BD"
        arrDone(65) = "False"
        arrValues(65) = "08C"
        arrDone(66) = "False"
        arrValues(66) = "289BD"
        arrDone(67) = "False"
        arrValues(67) = "0BCD"
        arrDone(68) = "True"
        arrValues(68) = "3"
        arrDone(69) = "False"
        arrValues(69) = "159C"
        arrDone(70) = "True"
        arrValues(70) = "A"
        arrDone(71) = "True"
        arrValues(71) = "F"
        arrDone(72) = "False"
        arrValues(72) = "78BC"
        arrDone(73) = "False"
        arrValues(73) = "278CD"
        arrDone(74) = "False"
        arrValues(74) = "0127D"
        arrDone(75) = "True"
        arrValues(75) = "4"
        arrDone(76) = "True"
        arrValues(76) = "E"
        arrDone(77) = "True"
        arrValues(77) = "6"
        arrDone(78) = "False"
        arrValues(78) = "079C"
        arrDone(79) = "False"
        arrValues(79) = "019C"
        arrDone(80) = "False"
        arrValues(80) = "8BD"
        arrDone(81) = "False"
        arrValues(81) = "68CE"
        arrDone(82) = "True"
        arrValues(82) = "5"
        arrDone(83) = "True"
        arrValues(83) = "F"
        arrDone(84) = "True"
        arrValues(84) = "0"
        arrDone(85) = "False"
        arrValues(85) = "146CE"
        arrDone(86) = "False"
        arrValues(86) = "14678C"
        arrDone(87) = "False"
        arrValues(87) = "1468BCE"
        arrDone(88) = "False"
        arrValues(88) = "678BC"
        arrDone(89) = "False"
        arrValues(89) = "678CD"
        arrDone(90) = "False"
        arrValues(90) = "17D"
        arrDone(91) = "True"
        arrValues(91) = "9"
        arrDone(92) = "False"
        arrValues(92) = "147"
        arrDone(93) = "True"
        arrValues(93) = "A"
        arrDone(94) = "True"
        arrValues(94) = "3"
        arrDone(95) = "True"
        arrValues(95) = "2"
        arrDone(96) = "True"
        arrValues(96) = "4"
        arrDone(97) = "True"
        arrValues(97) = "1"
        arrDone(98) = "False"
        arrValues(98) = "269AD"
        arrDone(99) = "True"
        arrValues(99) = "7"
        arrDone(100) = "False"
        arrValues(100) = "69CD"
        arrDone(101) = "False"
        arrValues(101) = "69CE"
        arrDone(102) = "False"
        arrValues(102) = "69C"
        arrDone(103) = "False"
        arrValues(103) = "6CE"
        arrDone(104) = "True"
        arrValues(104) = "5"
        arrDone(105) = "False"
        arrValues(105) = "236CD"
        arrDone(106) = "False"
        arrValues(106) = "023DF"
        arrDone(107) = "False"
        arrValues(107) = "023CDF"
        arrDone(108) = "True"
        arrValues(108) = "B"
        arrDone(109) = "True"
        arrValues(109) = "8"
        arrDone(110) = "False"
        arrValues(110) = "09C"
        arrDone(111) = "False"
        arrValues(111) = "09CF"
        arrDone(112) = "False"
        arrValues(112) = "89B"
        arrDone(113) = "False"
        arrValues(113) = "068C"
        arrDone(114) = "True"
        arrValues(114) = "3"
        arrDone(115) = "False"
        arrValues(115) = "06BC"
        arrDone(116) = "False"
        arrValues(116) = "16789BC"
        arrDone(117) = "False"
        arrValues(117) = "1469C"
        arrDone(118) = "False"
        arrValues(118) = "146789C"
        arrDone(119) = "True"
        arrValues(119) = "2"
        arrDone(120) = "True"
        arrValues(120) = "E"
        arrDone(121) = "True"
        arrValues(121) = "A"
        arrDone(122) = "False"
        arrValues(122) = "017F"
        arrDone(123) = "False"
        arrValues(123) = "0178CF"
        arrDone(124) = "False"
        arrValues(124) = "01479F"
        arrDone(125) = "False"
        arrValues(125) = "049CF"
        arrDone(126) = "True"
        arrValues(126) = "5"
        arrDone(127) = "True"
        arrValues(127) = "D"
        arrDone(128) = "True"
        arrValues(128) = "C"
        arrDone(129) = "True"
        arrValues(129) = "2"
        arrDone(130) = "False"
        arrValues(130) = "16DF"
        arrDone(131) = "False"
        arrValues(131) = "1356D"
        arrDone(132) = "False"
        arrValues(132) = "1689F"
        arrDone(133) = "False"
        arrValues(133) = "1469AF"
        arrDone(134) = "True"
        arrValues(134) = "B"
        arrDone(135) = "True"
        arrValues(135) = "7"
        arrDone(136) = "True"
        arrValues(136) = "0"
        arrDone(137) = "False"
        arrValues(137) = "34589D"
        arrDone(138) = "False"
        arrValues(138) = "135D"
        arrDone(139) = "False"
        arrValues(139) = "1358D"
        arrDone(140) = "False"
        arrValues(140) = "14589AF"
        arrDone(141) = "True"
        arrValues(141) = "E"
        arrDone(142) = "False"
        arrValues(142) = "49D"
        arrDone(143) = "False"
        arrValues(143) = "189AF"
        arrDone(144) = "False"
        arrValues(144) = "157F"
        arrDone(145) = "False"
        arrValues(145) = "5F"
        arrDone(146) = "True"
        arrValues(146) = "E"
        arrDone(147) = "True"
        arrValues(147) = "9"
        arrDone(148) = "False"
        arrValues(148) = "128CF"
        arrDone(149) = "False"
        arrValues(149) = "124ACF"
        arrDone(150) = "False"
        arrValues(150) = "01248CF"
        arrDone(151) = "True"
        arrValues(151) = "D"
        arrDone(152) = "False"
        arrValues(152) = "478C"
        arrDone(153) = "False"
        arrValues(153) = "24578C"
        arrDone(154) = "False"
        arrValues(154) = "1257"
        arrDone(155) = "False"
        arrValues(155) = "12578C"
        arrDone(156) = "True"
        arrValues(156) = "3"
        arrDone(157) = "False"
        arrValues(157) = "0245F"
        arrDone(158) = "True"
        arrValues(158) = "6"
        arrDone(159) = "True"
        arrValues(159) = "B"
        arrDone(160) = "True"
        arrValues(160) = "A"
        arrDone(161) = "True"
        arrValues(161) = "B"
        arrDone(162) = "True"
        arrValues(162) = "0"
        arrDone(163) = "False"
        arrValues(163) = "135D"
        arrDone(164) = "True"
        arrValues(164) = "E"
        arrDone(165) = "False"
        arrValues(165) = "1249F"
        arrDone(166) = "False"
        arrValues(166) = "123489F"
        arrDone(167) = "False"
        arrValues(167) = "1348"
        arrDone(168) = "False"
        arrValues(168) = "3489"
        arrDone(169) = "False"
        arrValues(169) = "234589D"
        arrDone(170) = "False"
        arrValues(170) = "1235D"
        arrDone(171) = "True"
        arrValues(171) = "6"
        arrDone(172) = "True"
        arrValues(172) = "C"
        arrDone(173) = "True"
        arrValues(173) = "7"
        arrDone(174) = "False"
        arrValues(174) = "249D"
        arrDone(175) = "False"
        arrValues(175) = "189F"
        arrDone(176) = "False"
        arrValues(176) = "137D"
        arrDone(177) = "False"
        arrValues(177) = "36"
        arrDone(178) = "True"
        arrValues(178) = "4"
        arrDone(179) = "True"
        arrValues(179) = "8"
        arrDone(180) = "True"
        arrValues(180) = "5"
        arrDone(181) = "False"
        arrValues(181) = "1269C"
        arrDone(182) = "False"
        arrValues(182) = "012369C"
        arrDone(183) = "False"
        arrValues(183) = "136C"
        arrDone(184) = "True"
        arrValues(184) = "F"
        arrDone(185) = "True"
        arrValues(185) = "B"
        arrDone(186) = "False"
        arrValues(186) = "1237DE"
        arrDone(187) = "True"
        arrValues(187) = "A"
        arrDone(188) = "False"
        arrValues(188) = "019"
        arrDone(189) = "False"
        arrValues(189) = "029D"
        arrDone(190) = "False"
        arrValues(190) = "029D"
        arrDone(191) = "False"
        arrValues(191) = "019"
        arrDone(192) = "False"
        arrValues(192) = "135F"
        arrDone(193) = "True"
        arrValues(193) = "9"
        arrDone(194) = "False"
        arrValues(194) = "1F"
        arrDone(195) = "True"
        arrValues(195) = "4"
        arrDone(196) = "False"
        arrValues(196) = "1267CF"
        arrDone(197) = "True"
        arrValues(197) = "B"
        arrDone(198) = "False"
        arrValues(198) = "12367CF"
        arrDone(199) = "False"
        arrValues(199) = "136CE"
        arrDone(200) = "False"
        arrValues(200) = "367C"
        arrDone(201) = "True"
        arrValues(201) = "0"
        arrDone(202) = "True"
        arrValues(202) = "8"
        arrDone(203) = "False"
        arrValues(203) = "2357CF"
        arrDone(204) = "False"
        arrValues(204) = "7F"
        arrDone(205) = "False"
        arrValues(205) = "23CDF"
        arrDone(206) = "True"
        arrValues(206) = "A"
        arrDone(207) = "False"
        arrValues(207) = "36CF"
        arrDone(208) = "True"
        arrValues(208) = "0"
        arrDone(209) = "True"
        arrValues(209) = "7"
        arrDone(210) = "False"
        arrValues(210) = "8ABF"
        arrDone(211) = "True"
        arrValues(211) = "2"
        arrDone(212) = "True"
        arrValues(212) = "4"
        arrDone(213) = "False"
        arrValues(213) = "6CEF"
        arrDone(214) = "False"
        arrValues(214) = "36CF"
        arrDone(215) = "True"
        arrValues(215) = "9"
        arrDone(216) = "True"
        arrValues(216) = "D"
        arrDone(217) = "False"
        arrValues(217) = "36CE"
        arrDone(218) = "False"
        arrValues(218) = "3AEF"
        arrDone(219) = "False"
        arrValues(219) = "3CF"
        arrDone(220) = "False"
        arrValues(220) = "8F"
        arrDone(221) = "True"
        arrValues(221) = "1"
        arrDone(222) = "False"
        arrValues(222) = "BCE"
        arrDone(223) = "True"
        arrValues(223) = "5"
        arrDone(224) = "False"
        arrValues(224) = "138BF"
        arrDone(225) = "False"
        arrValues(225) = "38CF"
        arrDone(226) = "False"
        arrValues(226) = "18BF"
        arrDone(227) = "True"
        arrValues(227) = "E"
        arrDone(228) = "True"
        arrValues(228) = "A"
        arrDone(229) = "True"
        arrValues(229) = "D"
        arrDone(230) = "True"
        arrValues(230) = "5"
        arrDone(231) = "False"
        arrValues(231) = "136C"
        arrDone(232) = "False"
        arrValues(232) = "3467C"
        arrDone(233) = "False"
        arrValues(233) = "23467C"
        arrDone(234) = "True"
        arrValues(234) = "9"
        arrDone(235) = "False"
        arrValues(235) = "237CF"
        arrDone(236) = "False"
        arrValues(236) = "078F"
        arrDone(237) = "False"
        arrValues(237) = "023BCF"
        arrDone(238) = "False"
        arrValues(238) = "027BC"
        arrDone(239) = "False"
        arrValues(239) = "0368CF"
        arrDone(240) = "True"
        arrValues(240) = "6"
        arrDone(241) = "True"
        arrValues(241) = "D"
        arrDone(242) = "False"
        arrValues(242) = "AF"
        arrDone(243) = "False"
        arrValues(243) = "35AC"
        arrDone(244) = "False"
        arrValues(244) = "27CF"
        arrDone(245) = "True"
        arrValues(245) = "8"
        arrDone(246) = "False"
        arrValues(246) = "237CF"
        arrDone(247) = "True"
        arrValues(247) = "0"
        arrDone(248) = "False"
        arrValues(248) = "37AC"
        arrDone(249) = "True"
        arrValues(249) = "1"
        arrDone(250) = "True"
        arrValues(250) = "B"
        arrDone(251) = "False"
        arrValues(251) = "2357CF"
        arrDone(252) = "False"
        arrValues(252) = "79F"
        arrDone(253) = "False"
        arrValues(253) = "239CF"
        arrDone(254) = "False"
        arrValues(254) = "279CE"
        arrDone(255) = "True"
        arrValues(255) = "4"

        TextBox3.Text = "Col3 has 8B duple, so R1C2 = 8 Col 3 has B so remove B from R1C3 And R2C3 Row5 has 17, 79, And 19 tuple, so remove 9 from R5C1"
        TextBox3.Text = "Col3 has 8B duple, so R1C2 = 8 Col 3 has B so remove B from R1C3 And R2C3 Row5 has 17, 79, And 19 tuple, so remove 9 from R5C1"
        RefreshPuzzle()

    End Sub

    Private Sub ButtonLoadSolution2_Click_1(sender As Object, e As EventArgs) Handles ButtonLoadSolution2.Click
        arrDone(0) = "True"
        arrValues(0) = "2"
        arrDone(1) = "False"
        arrValues(1) = "0568A"
        arrDone(2) = "False"
        arrValues(2) = "6789AB"
        arrDone(3) = "False"
        arrValues(3) = "056AB"
        arrDone(4) = "False"
        arrValues(4) = "689B"
        arrDone(5) = "True"
        arrValues(5) = "3"
        arrDone(6) = "True"
        arrValues(6) = "D"
        arrDone(7) = "False"
        arrValues(7) = "468AB"
        arrDone(8) = "True"
        arrValues(8) = "1"
        arrDone(9) = "False"
        arrValues(9) = "5789"
        arrDone(10) = "True"
        arrValues(10) = "C"
        arrDone(11) = "False"
        arrValues(11) = "0578"
        arrDone(12) = "False"
        arrValues(12) = "0459A"
        arrDone(13) = "False"
        arrValues(13) = "0459B"
        arrDone(14) = "True"
        arrValues(14) = "F"
        arrDone(15) = "True"
        arrValues(15) = "E"
        arrDone(16) = "False"
        arrValues(16) = "13589BD"
        arrDone(17) = "False"
        arrValues(17) = "0358A"
        arrDone(18) = "False"
        arrValues(18) = "189ABD"
        arrDone(19) = "False"
        arrValues(19) = "0135ABD"
        arrDone(20) = "False"
        arrValues(20) = "1289BC"
        arrDone(21) = "True"
        arrValues(21) = "7"
        arrDone(22) = "False"
        arrValues(22) = "1289C"
        arrDone(23) = "False"
        arrValues(23) = "18ABC"
        arrDone(24) = "False"
        arrValues(24) = "389A"
        arrDone(25) = "True"
        arrValues(25) = "F"
        arrDone(26) = "True"
        arrValues(26) = "4"
        arrDone(27) = "True"
        arrValues(27) = "E"
        arrDone(28) = "True"
        arrValues(28) = "6"
        arrDone(29) = "False"
        arrValues(29) = "0359BC"
        arrDone(30) = "False"
        arrValues(30) = "09BC"
        arrDone(31) = "False"
        arrValues(31) = "039AC"
        arrDone(32) = "True"
        arrValues(32) = "E"
        arrDone(33) = "False"
        arrValues(33) = "0368AF"
        arrDone(34) = "True"
        arrValues(34) = "C"
        arrDone(35) = "False"
        arrValues(35) = "036A"
        arrDone(36) = "False"
        arrValues(36) = "689F"
        arrDone(37) = "False"
        arrValues(37) = "469AF"
        arrDone(38) = "False"
        arrValues(38) = "4689F"
        arrDone(39) = "True"
        arrValues(39) = "5"
        arrDone(40) = "True"
        arrValues(40) = "2"
        arrDone(41) = "False"
        arrValues(41) = "389"
        arrDone(42) = "False"
        arrValues(42) = "03A"
        arrDone(43) = "True"
        arrValues(43) = "B"
        arrDone(44) = "True"
        arrValues(44) = "D"
        arrDone(45) = "False"
        arrValues(45) = "0349"
        arrDone(46) = "True"
        arrValues(46) = "1"
        arrDone(47) = "True"
        arrValues(47) = "7"
        arrDone(48) = "False"
        arrValues(48) = "13579BDF"
        arrDone(49) = "True"
        arrValues(49) = "4"
        arrDone(50) = "False"
        arrValues(50) = "179ABDF"
        arrDone(51) = "False"
        arrValues(51) = "135ABD"
        arrDone(52) = "False"
        arrValues(52) = "19BCF"
        arrDone(53) = "True"
        arrValues(53) = "0"
        arrDone(54) = "True"
        arrValues(54) = "E"
        arrDone(55) = "False"
        arrValues(55) = "1ABC"
        arrDone(56) = "False"
        arrValues(56) = "379A"
        arrDone(57) = "False"
        arrValues(57) = "3579D"
        arrDone(58) = "True"
        arrValues(58) = "6"
        arrDone(59) = "False"
        arrValues(59) = "357D"
        arrDone(60) = "True"
        arrValues(60) = "2"
        arrDone(61) = "False"
        arrValues(61) = "359BC"
        arrDone(62) = "True"
        arrValues(62) = "8"
        arrDone(63) = "False"
        arrValues(63) = "39AC"
        arrDone(64) = "False"
        arrValues(64) = "89BD"
        arrDone(65) = "False"
        arrValues(65) = "08C"
        arrDone(66) = "False"
        arrValues(66) = "289BD"
        arrDone(67) = "False"
        arrValues(67) = "0BCD"
        arrDone(68) = "True"
        arrValues(68) = "3"
        arrDone(69) = "False"
        arrValues(69) = "159C"
        arrDone(70) = "True"
        arrValues(70) = "A"
        arrDone(71) = "True"
        arrValues(71) = "F"
        arrDone(72) = "False"
        arrValues(72) = "78BC"
        arrDone(73) = "False"
        arrValues(73) = "278CD"
        arrDone(74) = "False"
        arrValues(74) = "0127D"
        arrDone(75) = "True"
        arrValues(75) = "4"
        arrDone(76) = "True"
        arrValues(76) = "E"
        arrDone(77) = "True"
        arrValues(77) = "6"
        arrDone(78) = "False"
        arrValues(78) = "079C"
        arrDone(79) = "False"
        arrValues(79) = "019C"
        arrDone(80) = "False"
        arrValues(80) = "8BD"
        arrDone(81) = "False"
        arrValues(81) = "68CE"
        arrDone(82) = "True"
        arrValues(82) = "5"
        arrDone(83) = "True"
        arrValues(83) = "F"
        arrDone(84) = "True"
        arrValues(84) = "0"
        arrDone(85) = "False"
        arrValues(85) = "146CE"
        arrDone(86) = "False"
        arrValues(86) = "14678C"
        arrDone(87) = "False"
        arrValues(87) = "1468BCE"
        arrDone(88) = "False"
        arrValues(88) = "678BC"
        arrDone(89) = "False"
        arrValues(89) = "678CD"
        arrDone(90) = "False"
        arrValues(90) = "17D"
        arrDone(91) = "True"
        arrValues(91) = "9"
        arrDone(92) = "False"
        arrValues(92) = "147"
        arrDone(93) = "True"
        arrValues(93) = "A"
        arrDone(94) = "True"
        arrValues(94) = "3"
        arrDone(95) = "True"
        arrValues(95) = "2"
        arrDone(96) = "True"
        arrValues(96) = "4"
        arrDone(97) = "True"
        arrValues(97) = "1"
        arrDone(98) = "False"
        arrValues(98) = "269AD"
        arrDone(99) = "True"
        arrValues(99) = "7"
        arrDone(100) = "False"
        arrValues(100) = "69CD"
        arrDone(101) = "False"
        arrValues(101) = "69CE"
        arrDone(102) = "False"
        arrValues(102) = "69C"
        arrDone(103) = "False"
        arrValues(103) = "6CE"
        arrDone(104) = "True"
        arrValues(104) = "5"
        arrDone(105) = "False"
        arrValues(105) = "236CD"
        arrDone(106) = "False"
        arrValues(106) = "023DF"
        arrDone(107) = "False"
        arrValues(107) = "023CDF"
        arrDone(108) = "True"
        arrValues(108) = "B"
        arrDone(109) = "True"
        arrValues(109) = "8"
        arrDone(110) = "False"
        arrValues(110) = "09C"
        arrDone(111) = "False"
        arrValues(111) = "09CF"
        arrDone(112) = "False"
        arrValues(112) = "89B"
        arrDone(113) = "False"
        arrValues(113) = "068C"
        arrDone(114) = "True"
        arrValues(114) = "3"
        arrDone(115) = "False"
        arrValues(115) = "06BC"
        arrDone(116) = "False"
        arrValues(116) = "16789BC"
        arrDone(117) = "False"
        arrValues(117) = "1469C"
        arrDone(118) = "False"
        arrValues(118) = "146789C"
        arrDone(119) = "True"
        arrValues(119) = "2"
        arrDone(120) = "True"
        arrValues(120) = "E"
        arrDone(121) = "True"
        arrValues(121) = "A"
        arrDone(122) = "False"
        arrValues(122) = "017F"
        arrDone(123) = "False"
        arrValues(123) = "0178CF"
        arrDone(124) = "False"
        arrValues(124) = "01479F"
        arrDone(125) = "False"
        arrValues(125) = "049CF"
        arrDone(126) = "True"
        arrValues(126) = "5"
        arrDone(127) = "True"
        arrValues(127) = "D"
        arrDone(128) = "True"
        arrValues(128) = "C"
        arrDone(129) = "True"
        arrValues(129) = "2"
        arrDone(130) = "False"
        arrValues(130) = "16DF"
        arrDone(131) = "False"
        arrValues(131) = "1356D"
        arrDone(132) = "False"
        arrValues(132) = "1689F"
        arrDone(133) = "False"
        arrValues(133) = "1469AF"
        arrDone(134) = "True"
        arrValues(134) = "B"
        arrDone(135) = "True"
        arrValues(135) = "7"
        arrDone(136) = "True"
        arrValues(136) = "0"
        arrDone(137) = "False"
        arrValues(137) = "34589D"
        arrDone(138) = "False"
        arrValues(138) = "135D"
        arrDone(139) = "False"
        arrValues(139) = "1358D"
        arrDone(140) = "False"
        arrValues(140) = "14589AF"
        arrDone(141) = "True"
        arrValues(141) = "E"
        arrDone(142) = "False"
        arrValues(142) = "49D"
        arrDone(143) = "False"
        arrValues(143) = "189AF"
        arrDone(144) = "False"
        arrValues(144) = "157F"
        arrDone(145) = "False"
        arrValues(145) = "5F"
        arrDone(146) = "True"
        arrValues(146) = "E"
        arrDone(147) = "True"
        arrValues(147) = "9"
        arrDone(148) = "False"
        arrValues(148) = "128CF"
        arrDone(149) = "False"
        arrValues(149) = "124ACF"
        arrDone(150) = "False"
        arrValues(150) = "01248CF"
        arrDone(151) = "True"
        arrValues(151) = "D"
        arrDone(152) = "False"
        arrValues(152) = "478C"
        arrDone(153) = "False"
        arrValues(153) = "24578C"
        arrDone(154) = "False"
        arrValues(154) = "1257"
        arrDone(155) = "False"
        arrValues(155) = "12578C"
        arrDone(156) = "True"
        arrValues(156) = "3"
        arrDone(157) = "False"
        arrValues(157) = "0245F"
        arrDone(158) = "True"
        arrValues(158) = "6"
        arrDone(159) = "True"
        arrValues(159) = "B"
        arrDone(160) = "True"
        arrValues(160) = "A"
        arrDone(161) = "True"
        arrValues(161) = "B"
        arrDone(162) = "True"
        arrValues(162) = "0"
        arrDone(163) = "False"
        arrValues(163) = "135D"
        arrDone(164) = "True"
        arrValues(164) = "E"
        arrDone(165) = "False"
        arrValues(165) = "1249F"
        arrDone(166) = "False"
        arrValues(166) = "123489F"
        arrDone(167) = "False"
        arrValues(167) = "1348"
        arrDone(168) = "False"
        arrValues(168) = "3489"
        arrDone(169) = "False"
        arrValues(169) = "234589D"
        arrDone(170) = "False"
        arrValues(170) = "1235D"
        arrDone(171) = "True"
        arrValues(171) = "6"
        arrDone(172) = "True"
        arrValues(172) = "C"
        arrDone(173) = "True"
        arrValues(173) = "7"
        arrDone(174) = "False"
        arrValues(174) = "249D"
        arrDone(175) = "False"
        arrValues(175) = "189F"
        arrDone(176) = "False"
        arrValues(176) = "137D"
        arrDone(177) = "False"
        arrValues(177) = "36"
        arrDone(178) = "True"
        arrValues(178) = "4"
        arrDone(179) = "True"
        arrValues(179) = "8"
        arrDone(180) = "True"
        arrValues(180) = "5"
        arrDone(181) = "False"
        arrValues(181) = "1269C"
        arrDone(182) = "False"
        arrValues(182) = "012369C"
        arrDone(183) = "False"
        arrValues(183) = "136C"
        arrDone(184) = "True"
        arrValues(184) = "F"
        arrDone(185) = "True"
        arrValues(185) = "B"
        arrDone(186) = "False"
        arrValues(186) = "1237DE"
        arrDone(187) = "True"
        arrValues(187) = "A"
        arrDone(188) = "False"
        arrValues(188) = "019"
        arrDone(189) = "False"
        arrValues(189) = "029D"
        arrDone(190) = "False"
        arrValues(190) = "029D"
        arrDone(191) = "False"
        arrValues(191) = "019"
        arrDone(192) = "False"
        arrValues(192) = "135F"
        arrDone(193) = "True"
        arrValues(193) = "9"
        arrDone(194) = "False"
        arrValues(194) = "1F"
        arrDone(195) = "True"
        arrValues(195) = "4"
        arrDone(196) = "False"
        arrValues(196) = "1267CF"
        arrDone(197) = "True"
        arrValues(197) = "B"
        arrDone(198) = "False"
        arrValues(198) = "12367CF"
        arrDone(199) = "False"
        arrValues(199) = "136CE"
        arrDone(200) = "False"
        arrValues(200) = "367C"
        arrDone(201) = "True"
        arrValues(201) = "0"
        arrDone(202) = "True"
        arrValues(202) = "8"
        arrDone(203) = "False"
        arrValues(203) = "2357CF"
        arrDone(204) = "False"
        arrValues(204) = "7F"
        arrDone(205) = "False"
        arrValues(205) = "23CDF"
        arrDone(206) = "True"
        arrValues(206) = "A"
        arrDone(207) = "False"
        arrValues(207) = "36CF"
        arrDone(208) = "True"
        arrValues(208) = "0"
        arrDone(209) = "True"
        arrValues(209) = "7"
        arrDone(210) = "False"
        arrValues(210) = "8ABF"
        arrDone(211) = "True"
        arrValues(211) = "2"
        arrDone(212) = "True"
        arrValues(212) = "4"
        arrDone(213) = "False"
        arrValues(213) = "6CEF"
        arrDone(214) = "False"
        arrValues(214) = "36CF"
        arrDone(215) = "True"
        arrValues(215) = "9"
        arrDone(216) = "True"
        arrValues(216) = "D"
        arrDone(217) = "False"
        arrValues(217) = "36CE"
        arrDone(218) = "False"
        arrValues(218) = "3AEF"
        arrDone(219) = "False"
        arrValues(219) = "3CF"
        arrDone(220) = "False"
        arrValues(220) = "8F"
        arrDone(221) = "True"
        arrValues(221) = "1"
        arrDone(222) = "False"
        arrValues(222) = "BCE"
        arrDone(223) = "True"
        arrValues(223) = "5"
        arrDone(224) = "False"
        arrValues(224) = "138BF"
        arrDone(225) = "False"
        arrValues(225) = "38CF"
        arrDone(226) = "False"
        arrValues(226) = "18BF"
        arrDone(227) = "True"
        arrValues(227) = "E"
        arrDone(228) = "True"
        arrValues(228) = "A"
        arrDone(229) = "True"
        arrValues(229) = "D"
        arrDone(230) = "True"
        arrValues(230) = "5"
        arrDone(231) = "False"
        arrValues(231) = "136C"
        arrDone(232) = "False"
        arrValues(232) = "3467C"
        arrDone(233) = "False"
        arrValues(233) = "23467C"
        arrDone(234) = "True"
        arrValues(234) = "9"
        arrDone(235) = "False"
        arrValues(235) = "237CF"
        arrDone(236) = "False"
        arrValues(236) = "078F"
        arrDone(237) = "False"
        arrValues(237) = "023BCF"
        arrDone(238) = "False"
        arrValues(238) = "027BC"
        arrDone(239) = "False"
        arrValues(239) = "0368CF"
        arrDone(240) = "True"
        arrValues(240) = "6"
        arrDone(241) = "True"
        arrValues(241) = "D"
        arrDone(242) = "False"
        arrValues(242) = "AF"
        arrDone(243) = "False"
        arrValues(243) = "35AC"
        arrDone(244) = "False"
        arrValues(244) = "27CF"
        arrDone(245) = "True"
        arrValues(245) = "8"
        arrDone(246) = "False"
        arrValues(246) = "237CF"
        arrDone(247) = "True"
        arrValues(247) = "0"
        arrDone(248) = "False"
        arrValues(248) = "37AC"
        arrDone(249) = "True"
        arrValues(249) = "1"
        arrDone(250) = "True"
        arrValues(250) = "B"
        arrDone(251) = "False"
        arrValues(251) = "2357CF"
        arrDone(252) = "False"
        arrValues(252) = "79F"
        arrDone(253) = "False"
        arrValues(253) = "239CF"
        arrDone(254) = "False"
        arrValues(254) = "279CE"
        arrDone(255) = "True"
        arrValues(255) = "4"

        RefreshPuzzle()

    End Sub
    Private Sub ButtonLoadSolution3_Click_1(sender As Object, e As EventArgs) Handles ButtonLoadSolution3.Click

        arrDone(0) = "True"
        arrValues(0) = "7"
        arrDone(1) = "False"
        arrValues(1) = "589CDF"
        arrDone(2) = "False"
        arrValues(2) = "89CDE"
        arrDone(3) = "True"
        arrValues(3) = "1"
        arrDone(4) = "False"
        arrValues(4) = "89ACDF"
        arrDone(5) = "False"
        arrValues(5) = "59AF"
        arrDone(6) = "False"
        arrValues(6) = "58ACD"
        arrDone(7) = "False"
        arrValues(7) = "589AF"
        arrDone(8) = "False"
        arrValues(8) = "8AB"
        arrDone(9) = "True"
        arrValues(9) = "2"
        arrDone(10) = "False"
        arrValues(10) = "58E"
        arrDone(11) = "True"
        arrValues(11) = "4"
        arrDone(12) = "False"
        arrValues(12) = "59AC"
        arrDone(13) = "True"
        arrValues(13) = "0"
        arrDone(14) = "True"
        arrValues(14) = "3"
        arrDone(15) = "True"
        arrValues(15) = "6"
        arrDone(16) = "False"
        arrValues(16) = "3456D"
        arrDone(17) = "False"
        arrValues(17) = "45689D"
        arrDone(18) = "True"
        arrValues(18) = "B"
        arrDone(19) = "True"
        arrValues(19) = "0"
        arrDone(20) = "True"
        arrValues(20) = "E"
        arrDone(21) = "True"
        arrValues(21) = "2"
        arrDone(22) = "False"
        arrValues(22) = "4568AD"
        arrDone(23) = "False"
        arrValues(23) = "3589A"
        arrDone(24) = "False"
        arrValues(24) = "678A"
        arrDone(25) = "False"
        arrValues(25) = "5678A"
        arrDone(26) = "True"
        arrValues(26) = "C"
        arrDone(27) = "True"
        arrValues(27) = "1"
        arrDone(28) = "True"
        arrValues(28) = "F"
        arrDone(29) = "False"
        arrValues(29) = "4579D"
        arrDone(30) = "False"
        arrValues(30) = "459A"
        arrDone(31) = "False"
        arrValues(31) = "47A"
        arrDone(32) = "True"
        arrValues(32) = "A"
        arrDone(33) = "False"
        arrValues(33) = "45689CD"
        arrDone(34) = "False"
        arrValues(34) = "689CDE"
        arrDone(35) = "False"
        arrValues(35) = "4569CD"
        arrDone(36) = "False"
        arrValues(36) = "0689CD"
        arrDone(37) = "False"
        arrValues(37) = "1459"
        arrDone(38) = "False"
        arrValues(38) = "014568CD"
        arrDone(39) = "False"
        arrValues(39) = "0589"
        arrDone(40) = "False"
        arrValues(40) = "678B"
        arrDone(41) = "True"
        arrValues(41) = "F"
        arrDone(42) = "True"
        arrValues(42) = "3"
        arrDone(43) = "False"
        arrValues(43) = "0678BE"
        arrDone(44) = "False"
        arrValues(44) = "4579C"
        arrDone(45) = "True"
        arrValues(45) = "2"
        arrDone(46) = "False"
        arrValues(46) = "1459BCE"
        arrDone(47) = "False"
        arrValues(47) = "147BE"
        arrDone(48) = "False"
        arrValues(48) = "3456CEF"
        arrDone(49) = "False"
        arrValues(49) = "2456CF"
        arrDone(50) = "False"
        arrValues(50) = "26CE"
        arrDone(51) = "False"
        arrValues(51) = "2456C"
        arrDone(52) = "False"
        arrValues(52) = "06ACF"
        arrDone(53) = "False"
        arrValues(53) = "1345AF"
        arrDone(54) = "True"
        arrValues(54) = "7"
        arrDone(55) = "True"
        arrValues(55) = "B"
        arrDone(56) = "True"
        arrValues(56) = "9"
        arrDone(57) = "False"
        arrValues(57) = "056AE"
        arrDone(58) = "True"
        arrValues(58) = "D"
        arrDone(59) = "False"
        arrValues(59) = "06E"
        arrDone(60) = "True"
        arrValues(60) = "8"
        arrDone(61) = "False"
        arrValues(61) = "145C"
        arrDone(62) = "False"
        arrValues(62) = "145ACE"
        arrDone(63) = "False"
        arrValues(63) = "14AE"
        arrDone(64) = "False"
        arrValues(64) = "056D"
        arrDone(65) = "False"
        arrValues(65) = "12569ABD"
        arrDone(66) = "True"
        arrValues(66) = "3"
        arrDone(67) = "False"
        arrValues(67) = "2569BD"
        arrDone(68) = "False"
        arrValues(68) = "079AB"
        arrDone(69) = "False"
        arrValues(69) = "4579AB"
        arrDone(70) = "False"
        arrValues(70) = "045AB"
        arrDone(71) = "False"
        arrValues(71) = "0579A"
        arrDone(72) = "True"
        arrValues(72) = "F"
        arrDone(73) = "False"
        arrValues(73) = "4567DE"
        arrDone(74) = "False"
        arrValues(74) = "4567E"
        arrDone(75) = "False"
        arrValues(75) = "67BDE"
        arrDone(76) = "False"
        arrValues(76) = "045"
        arrDone(77) = "True"
        arrValues(77) = "8"
        arrDone(78) = "False"
        arrValues(78) = "0145B"
        arrDone(79) = "True"
        arrValues(79) = "C"
        arrDone(80) = "False"
        arrValues(80) = "056"
        arrDone(81) = "True"
        arrValues(81) = "E"
        arrDone(82) = "False"
        arrValues(82) = "1679"
        arrDone(83) = "True"
        arrValues(83) = "F"
        arrDone(84) = "False"
        arrValues(84) = "0789B"
        arrDone(85) = "True"
        arrValues(85) = "D"
        arrDone(86) = "False"
        arrValues(86) = "0458B"
        arrDone(87) = "True"
        arrValues(87) = "C"
        arrDone(88) = "False"
        arrValues(88) = "34678B"
        arrDone(89) = "False"
        arrValues(89) = "45678"
        arrDone(90) = "False"
        arrValues(90) = "45678"
        arrDone(91) = "False"
        arrValues(91) = "3678B"
        arrDone(92) = "True"
        arrValues(92) = "2"
        arrDone(93) = "True"
        arrValues(93) = "A"
        arrDone(94) = "False"
        arrValues(94) = "0145B"
        arrDone(95) = "False"
        arrValues(95) = "0134B"
        arrDone(96) = "True"
        arrValues(96) = "8"
        arrDone(97) = "False"
        arrValues(97) = "5BC"
        arrDone(98) = "False"
        arrValues(98) = "7C"
        arrDone(99) = "False"
        arrValues(99) = "5BC"
        arrDone(100) = "True"
        arrValues(100) = "3"
        arrDone(101) = "False"
        arrValues(101) = "457BF"
        arrDone(102) = "True"
        arrValues(102) = "2"
        arrDone(103) = "True"
        arrValues(103) = "E"
        arrDone(104) = "True"
        arrValues(104) = "0"
        arrDone(105) = "True"
        arrValues(105) = "1"
        arrDone(106) = "True"
        arrValues(106) = "A"
        arrDone(107) = "False"
        arrValues(107) = "7BC"
        arrDone(108) = "True"
        arrValues(108) = "6"
        arrDone(109) = "False"
        arrValues(109) = "45"
        arrDone(110) = "True"
        arrValues(110) = "D"
        arrDone(111) = "True"
        arrValues(111) = "9"
        arrDone(112) = "False"
        arrValues(112) = "05CD"
        arrDone(113) = "False"
        arrValues(113) = "5ABCD"
        arrDone(114) = "True"
        arrValues(114) = "4"
        arrDone(115) = "False"
        arrValues(115) = "5BCD"
        arrDone(116) = "True"
        arrValues(116) = "1"
        arrDone(117) = "False"
        arrValues(117) = "5AB"
        arrDone(118) = "False"
        arrValues(118) = "058AB"
        arrDone(119) = "True"
        arrValues(119) = "6"
        arrDone(120) = "False"
        arrValues(120) = "38BC"
        arrDone(121) = "False"
        arrValues(121) = "58CD"
        arrDone(122) = "True"
        arrValues(122) = "9"
        arrDone(123) = "True"
        arrValues(123) = "2"
        arrDone(124) = "True"
        arrValues(124) = "E"
        arrDone(125) = "False"
        arrValues(125) = "35"
        arrDone(126) = "True"
        arrValues(126) = "7"
        arrDone(127) = "True"
        arrValues(127) = "F"
        arrDone(128) = "True"
        arrValues(128) = "B"
        arrDone(129) = "True"
        arrValues(129) = "0"
        arrDone(130) = "False"
        arrValues(130) = "CD"
        arrDone(131) = "True"
        arrValues(131) = "8"
        arrDone(132) = "True"
        arrValues(132) = "2"
        arrDone(133) = "True"
        arrValues(133) = "6"
        arrDone(134) = "False"
        arrValues(134) = "AD"
        arrDone(135) = "False"
        arrValues(135) = "379A"
        arrDone(136) = "True"
        arrValues(136) = "1"
        arrDone(137) = "False"
        arrValues(137) = "479C"
        arrDone(138) = "False"
        arrValues(138) = "47"
        arrDone(139) = "True"
        arrValues(139) = "5"
        arrDone(140) = "False"
        arrValues(140) = "3479AC"
        arrDone(141) = "True"
        arrValues(141) = "E"
        arrDone(142) = "False"
        arrValues(142) = "49ACF"
        arrDone(143) = "False"
        arrValues(143) = "347A"
        arrDone(144) = "True"
        arrValues(144) = "2"
        arrDone(145) = "True"
        arrValues(145) = "7"
        arrDone(146) = "False"
        arrValues(146) = "16E"
        arrDone(147) = "True"
        arrValues(147) = "3"
        arrDone(148) = "False"
        arrValues(148) = "089"
        arrDone(149) = "True"
        arrValues(149) = "C"
        arrDone(150) = "True"
        arrValues(150) = "F"
        arrDone(151) = "True"
        arrValues(151) = "4"
        arrDone(152) = "True"
        arrValues(152) = "D"
        arrDone(153) = "True"
        arrValues(153) = "B"
        arrDone(154) = "False"
        arrValues(154) = "068"
        arrDone(155) = "True"
        arrValues(155) = "A"
        arrDone(156) = "False"
        arrValues(156) = "09"
        arrDone(157) = "False"
        arrValues(157) = "169"
        arrDone(158) = "False"
        arrValues(158) = "01689"
        arrDone(159) = "True"
        arrValues(159) = "5"
        arrDone(160) = "False"
        arrValues(160) = "46C"
        arrDone(161) = "False"
        arrValues(161) = "146C"
        arrDone(162) = "True"
        arrValues(162) = "5"
        arrDone(163) = "True"
        arrValues(163) = "A"
        arrDone(164) = "False"
        arrValues(164) = "0789B"
        arrDone(165) = "False"
        arrValues(165) = "379B"
        arrDone(166) = "False"
        arrValues(166) = "08B"
        arrDone(167) = "False"
        arrValues(167) = "03789"
        arrDone(168) = "True"
        arrValues(168) = "E"
        arrDone(169) = "False"
        arrValues(169) = "046789C"
        arrDone(170) = "True"
        arrValues(170) = "F"
        arrDone(171) = "False"
        arrValues(171) = "03678C"
        arrDone(172) = "True"
        arrValues(172) = "D"
        arrDone(173) = "False"
        arrValues(173) = "134679C"
        arrDone(174) = "True"
        arrValues(174) = "2"
        arrDone(175) = "False"
        arrValues(175) = "013478"
        arrDone(176) = "True"
        arrValues(176) = "9"
        arrDone(177) = "False"
        arrValues(177) = "46CD"
        arrDone(178) = "True"
        arrValues(178) = "F"
        arrDone(179) = "False"
        arrValues(179) = "46CD"
        arrDone(180) = "False"
        arrValues(180) = "078AD"
        arrDone(181) = "False"
        arrValues(181) = "357A"
        arrDone(182) = "False"
        arrValues(182) = "058ADE"
        arrDone(183) = "True"
        arrValues(183) = "1"
        arrDone(184) = "False"
        arrValues(184) = "234678C"
        arrDone(185) = "False"
        arrValues(185) = "04678C"
        arrDone(186) = "False"
        arrValues(186) = "024678"
        arrDone(187) = "False"
        arrValues(187) = "03678C"
        arrDone(188) = "False"
        arrValues(188) = "0347AC"
        arrDone(189) = "True"
        arrValues(189) = "B"
        arrDone(190) = "False"
        arrValues(190) = "0468AC"
        arrDone(191) = "False"
        arrValues(191) = "03478A"
        arrDone(192) = "False"
        arrValues(192) = "46CF"
        arrDone(193) = "False"
        arrValues(193) = "24689BCF"
        arrDone(194) = "False"
        arrValues(194) = "2689C"
        arrDone(195) = "True"
        arrValues(195) = "E"
        arrDone(196) = "False"
        arrValues(196) = "67ABCF"
        arrDone(197) = "True"
        arrValues(197) = "0"
        arrDone(198) = "False"
        arrValues(198) = "16ABC"
        arrDone(199) = "True"
        arrValues(199) = "D"
        arrDone(200) = "True"
        arrValues(200) = "5"
        arrDone(201) = "True"
        arrValues(201) = "3"
        arrDone(202) = "False"
        arrValues(202) = "124678"
        arrDone(203) = "False"
        arrValues(203) = "678CF"
        arrDone(204) = "False"
        arrValues(204) = "479AC"
        arrDone(205) = "False"
        arrValues(205) = "4679C"
        arrDone(206) = "False"
        arrValues(206) = "4689AC"
        arrDone(207) = "False"
        arrValues(207) = "478A"
        arrDone(208) = "False"
        arrValues(208) = "456CDF"
        arrDone(209) = "False"
        arrValues(209) = "45689BCDF"
        arrDone(210) = "True"
        arrValues(210) = "A"
        arrDone(211) = "False"
        arrValues(211) = "4569BCD"
        arrDone(212) = "False"
        arrValues(212) = "67BCF"
        arrDone(213) = "True"
        arrValues(213) = "E"
        arrDone(214) = "True"
        arrValues(214) = "3"
        arrDone(215) = "False"
        arrValues(215) = "7F"
        arrDone(216) = "False"
        arrValues(216) = "4678C"
        arrDone(217) = "False"
        arrValues(217) = "04678CD"
        arrDone(218) = "False"
        arrValues(218) = "014678"
        arrDone(219) = "False"
        arrValues(219) = "0678CDF"
        arrDone(220) = "False"
        arrValues(220) = "04579C"
        arrDone(221) = "False"
        arrValues(221) = "45679C"
        arrDone(222) = "False"
        arrValues(222) = "045689C"
        arrDone(223) = "True"
        arrValues(223) = "2"
        arrDone(224) = "False"
        arrValues(224) = "56CD"
        arrDone(225) = "False"
        arrValues(225) = "256CD"
        arrDone(226) = "False"
        arrValues(226) = "26CD"
        arrDone(227) = "True"
        arrValues(227) = "7"
        arrDone(228) = "True"
        arrValues(228) = "4"
        arrDone(229) = "True"
        arrValues(229) = "8"
        arrDone(230) = "False"
        arrValues(230) = "6AC"
        arrDone(231) = "False"
        arrValues(231) = "2A"
        arrDone(232) = "False"
        arrValues(232) = "26AC"
        arrDone(233) = "False"
        arrValues(233) = "06ACDE"
        arrDone(234) = "True"
        arrValues(234) = "B"
        arrDone(235) = "True"
        arrValues(235) = "9"
        arrDone(236) = "True"
        arrValues(236) = "1"
        arrDone(237) = "True"
        arrValues(237) = "F"
        arrDone(238) = "False"
        arrValues(238) = "056ACE"
        arrDone(239) = "False"
        arrValues(239) = "03AE"
        arrDone(240) = "True"
        arrValues(240) = "1"
        arrDone(241) = "True"
        arrValues(241) = "3"
        arrDone(242) = "True"
        arrValues(242) = "0"
        arrDone(243) = "False"
        arrValues(243) = "246C"
        arrDone(244) = "True"
        arrValues(244) = "5"
        arrDone(245) = "False"
        arrValues(245) = "7AF"
        arrDone(246) = "True"
        arrValues(246) = "9"
        arrDone(247) = "False"
        arrValues(247) = "27AF"
        arrDone(248) = "False"
        arrValues(248) = "24678AC"
        arrDone(249) = "False"
        arrValues(249) = "4678ACE"
        arrDone(250) = "False"
        arrValues(250) = "24678E"
        arrDone(251) = "False"
        arrValues(251) = "678CEF"
        arrDone(252) = "True"
        arrValues(252) = "B"
        arrDone(253) = "False"
        arrValues(253) = "467C"
        arrDone(254) = "False"
        arrValues(254) = "468ACE"
        arrDone(255) = "True"
        arrValues(255) = "D"

        RefreshPuzzle()

    End Sub

    Private Sub LoadSolution3A()

        arrDone(0) = "False"
        arrValues(0) = "158A"
        arrDone(1) = "False"
        arrValues(1) = "03568"
        arrDone(2) = "False"
        arrValues(2) = "0359AD"
        arrDone(3) = "False"
        arrValues(3) = "13569AD"
        arrDone(4) = "False"
        arrValues(4) = "135F"
        arrDone(5) = "True"
        arrValues(5) = "B"
        arrDone(6) = "True"
        arrValues(6) = "4"
        arrDone(7) = "True"
        arrValues(7) = "C"
        arrDone(8) = "False"
        arrValues(8) = "156D"
        arrDone(9) = "True"
        arrValues(9) = "E"
        arrDone(10) = "False"
        arrValues(10) = "016A"
        arrDone(11) = "True"
        arrValues(11) = "2"
        arrDone(12) = "False"
        arrValues(12) = "0135AD"
        arrDone(13) = "False"
        arrValues(13) = "0135ADF"
        arrDone(14) = "False"
        arrValues(14) = "013ADF"
        arrDone(15) = "True"
        arrValues(15) = "7"
        arrDone(16) = "False"
        arrValues(16) = "145C"
        arrDone(17) = "False"
        arrValues(17) = "03457E"
        arrDone(18) = "False"
        arrValues(18) = "03457CDE"
        arrDone(19) = "False"
        arrValues(19) = "13457CD"
        arrDone(20) = "False"
        arrValues(20) = "135"
        arrDone(21) = "False"
        arrValues(21) = "35D"
        arrDone(22) = "True"
        arrValues(22) = "A"
        arrDone(23) = "True"
        arrValues(23) = "2"
        arrDone(24) = "False"
        arrValues(24) = "157CD"
        arrDone(25) = "True"
        arrValues(25) = "F"
        arrDone(26) = "True"
        arrValues(26) = "B"
        arrDone(27) = "False"
        arrValues(27) = "1457D"
        arrDone(28) = "True"
        arrValues(28) = "8"
        arrDone(29) = "True"
        arrValues(29) = "9"
        arrDone(30) = "False"
        arrValues(30) = "0134DE"
        arrDone(31) = "True"
        arrValues(31) = "6"
        arrDone(32) = "False"
        arrValues(32) = "145AB"
        arrDone(33) = "False"
        arrValues(33) = "34567BE"
        arrDone(34) = "False"
        arrValues(34) = "34579ADE"
        arrDone(35) = "True"
        arrValues(35) = "2"
        arrDone(36) = "True"
        arrValues(36) = "0"
        arrDone(37) = "False"
        arrValues(37) = "356DF"
        arrDone(38) = "False"
        arrValues(38) = "356D"
        arrDone(39) = "False"
        arrValues(39) = "156F"
        arrDone(40) = "False"
        arrValues(40) = "1567D"
        arrDone(41) = "False"
        arrValues(41) = "679AD"
        arrDone(42) = "True"
        arrValues(42) = "8"
        arrDone(43) = "False"
        arrValues(43) = "145679AD"
        arrDone(44) = "False"
        arrValues(44) = "135ABDE"
        arrDone(45) = "True"
        arrValues(45) = "C"
        arrDone(46) = "False"
        arrValues(46) = "134ABDEF"
        arrDone(47) = "False"
        arrValues(47) = "1345F"
        arrDone(48) = "False"
        arrValues(48) = "145ABC"
        arrDone(49) = "False"
        arrValues(49) = "0456B"
        arrDone(50) = "False"
        arrValues(50) = "045ACD"
        arrDone(51) = "True"
        arrValues(51) = "F"
        arrDone(52) = "True"
        arrValues(52) = "9"
        arrDone(53) = "True"
        arrValues(53) = "8"
        arrDone(54) = "True"
        arrValues(54) = "7"
        arrDone(55) = "True"
        arrValues(55) = "E"
        arrDone(56) = "True"
        arrValues(56) = "3"
        arrDone(57) = "False"
        arrValues(57) = "6ACD"
        arrDone(58) = "False"
        arrValues(58) = "0146A"
        arrDone(59) = "False"
        arrValues(59) = "1456AD"
        arrDone(60) = "False"
        arrValues(60) = "0125ABD"
        arrDone(61) = "False"
        arrValues(61) = "01245ABD"
        arrDone(62) = "False"
        arrValues(62) = "014ABD"
        arrDone(63) = "False"
        arrValues(63) = "1245"
        arrDone(64) = "True"
        arrValues(64) = "6"
        arrDone(65) = "True"
        arrValues(65) = "C"
        arrDone(66) = "False"
        arrValues(66) = "357AE"
        arrDone(67) = "False"
        arrValues(67) = "1357A"
        arrDone(68) = "False"
        arrValues(68) = "358AE"
        arrDone(69) = "True"
        arrValues(69) = "4"
        arrDone(70) = "False"
        arrValues(70) = "3589"
        arrDone(71) = "True"
        arrValues(71) = "D"
        arrDone(72) = "True"
        arrValues(72) = "F"
        arrDone(73) = "True"
        arrValues(73) = "2"
        arrDone(74) = "False"
        arrValues(74) = "137AE"
        arrDone(75) = "False"
        arrValues(75) = "1578A"
        arrDone(76) = "False"
        arrValues(76) = "01357"
        arrDone(77) = "False"
        arrValues(77) = "01357"
        arrDone(78) = "False"
        arrValues(78) = "013"
        arrDone(79) = "True"
        arrValues(79) = "B"
        arrDone(80) = "False"
        arrValues(80) = "125B"
        arrDone(81) = "True"
        arrValues(81) = "F"
        arrDone(82) = "False"
        arrValues(82) = "235E"
        arrDone(83) = "False"
        arrValues(83) = "135B"
        arrDone(84) = "False"
        arrValues(84) = "358BE"
        arrDone(85) = "True"
        arrValues(85) = "7"
        arrDone(86) = "False"
        arrValues(86) = "358B"
        arrDone(87) = "True"
        arrValues(87) = "0"
        arrDone(88) = "True"
        arrValues(88) = "9"
        arrDone(89) = "False"
        arrValues(89) = "368CD"
        arrDone(90) = "False"
        arrValues(90) = "136E"
        arrDone(91) = "False"
        arrValues(91) = "1568BD"
        arrDone(92) = "True"
        arrValues(92) = "4"
        arrDone(93) = "False"
        arrValues(93) = "1235D"
        arrDone(94) = "False"
        arrValues(94) = "13CD"
        arrDone(95) = "True"
        arrValues(95) = "A"
        arrDone(96) = "False"
        arrValues(96) = "145AB"
        arrDone(97) = "True"
        arrValues(97) = "9"
        arrDone(98) = "True"
        arrValues(98) = "8"
        arrDone(99) = "False"
        arrValues(99) = "13457AB"
        arrDone(100) = "True"
        arrValues(100) = "2"
        arrDone(101) = "False"
        arrValues(101) = "35AE"
        arrDone(102) = "True"
        arrValues(102) = "C"
        arrDone(103) = "False"
        arrValues(103) = "5A"
        arrDone(104) = "False"
        arrValues(104) = "157BDE"
        arrDone(105) = "False"
        arrValues(105) = "37AD"
        arrDone(106) = "False"
        arrValues(106) = "137AE"
        arrDone(107) = "True"
        arrValues(107) = "0"
        arrDone(108) = "True"
        arrValues(108) = "F"
        arrDone(109) = "False"
        arrValues(109) = "1357D"
        arrDone(110) = "True"
        arrValues(110) = "6"
        arrDone(111) = "False"
        arrValues(111) = "135"
        arrDone(112) = "True"
        arrValues(112) = "0"
        arrDone(113) = "True"
        arrValues(113) = "D"
        arrDone(114) = "False"
        arrValues(114) = "2357A"
        arrDone(115) = "False"
        arrValues(115) = "357AB"
        arrDone(116) = "True"
        arrValues(116) = "6"
        arrDone(117) = "True"
        arrValues(117) = "1"
        arrDone(118) = "False"
        arrValues(118) = "35B"
        arrDone(119) = "False"
        arrValues(119) = "5AF"
        arrDone(120) = "False"
        arrValues(120) = "57BC"
        arrDone(121) = "True"
        arrValues(121) = "4"
        arrDone(122) = "False"
        arrValues(122) = "37A"
        arrDone(123) = "False"
        arrValues(123) = "57AB"
        arrDone(124) = "False"
        arrValues(124) = "2357C"
        arrDone(125) = "True"
        arrValues(125) = "8"
        arrDone(126) = "True"
        arrValues(126) = "9"
        arrDone(127) = "True"
        arrValues(127) = "E"
        arrDone(128) = "True"
        arrValues(128) = "3"
        arrDone(129) = "True"
        arrValues(129) = "A"
        arrDone(130) = "True"
        arrValues(130) = "6"
        arrDone(131) = "False"
        arrValues(131) = "457B"
        arrDone(132) = "False"
        arrValues(132) = "57E"
        arrDone(133) = "False"
        arrValues(133) = "59E"
        arrDone(134) = "True"
        arrValues(134) = "0"
        arrDone(135) = "False"
        arrValues(135) = "4579"
        arrDone(136) = "False"
        arrValues(136) = "17E"
        arrDone(137) = "False"
        arrValues(137) = "789"
        arrDone(138) = "True"
        arrValues(138) = "F"
        arrDone(139) = "True"
        arrValues(139) = "C"
        arrDone(140) = "False"
        arrValues(140) = "1B"
        arrDone(141) = "False"
        arrValues(141) = "14B"
        arrDone(142) = "True"
        arrValues(142) = "2"
        arrDone(143) = "True"
        arrValues(143) = "D"
        arrDone(144) = "False"
        arrValues(144) = "245CF"
        arrDone(145) = "True"
        arrValues(145) = "1"
        arrDone(146) = "False"
        arrValues(146) = "0245CF"
        arrDone(147) = "True"
        arrValues(147) = "8"
        arrDone(148) = "True"
        arrValues(148) = "D"
        arrDone(149) = "False"
        arrValues(149) = "2569AF"
        arrDone(150) = "False"
        arrValues(150) = "2569"
        arrDone(151) = "False"
        arrValues(151) = "4569AF"
        arrDone(152) = "False"
        arrValues(152) = "6"
        arrDone(153) = "True"
        arrValues(153) = "B"
        arrDone(154) = "False"
        arrValues(154) = "46A"
        arrDone(155) = "True"
        arrValues(155) = "3"
        arrDone(156) = "False"
        arrValues(156) = "AC"
        arrDone(157) = "True"
        arrValues(157) = "E"
        arrDone(158) = "True"
        arrValues(158) = "7"
        arrDone(159) = "False"
        arrValues(159) = "49CF"
        arrDone(160) = "True"
        arrValues(160) = "D"
        arrDone(161) = "False"
        arrValues(161) = "247"
        arrDone(162) = "False"
        arrValues(162) = "247CF"
        arrDone(163) = "True"
        arrValues(163) = "E"
        arrDone(164) = "False"
        arrValues(164) = "37ACF"
        arrDone(165) = "False"
        arrValues(165) = "2369AF"
        arrDone(166) = "False"
        arrValues(166) = "2369"
        arrDone(167) = "True"
        arrValues(167) = "B"
        arrDone(168) = "True"
        arrValues(168) = "0"
        arrDone(169) = "False"
        arrValues(169) = "679A"
        arrDone(170) = "True"
        arrValues(170) = "5"
        arrDone(171) = "False"
        arrValues(171) = "14679A"
        arrDone(172) = "False"
        arrValues(172) = "13AC"
        arrDone(173) = "False"
        arrValues(173) = "1346AF"
        arrDone(174) = "True"
        arrValues(174) = "8"
        arrDone(175) = "False"
        arrValues(175) = "1349CF"
        arrDone(176) = "True"
        arrValues(176) = "9"
        arrDone(177) = "False"
        arrValues(177) = "47B"
        arrDone(178) = "False"
        arrValues(178) = "47CF"
        arrDone(179) = "False"
        arrValues(179) = "47BC"
        arrDone(180) = "False"
        arrValues(180) = "37ACEF"
        arrDone(181) = "False"
        arrValues(181) = "36AEF"
        arrDone(182) = "True"
        arrValues(182) = "1"
        arrDone(183) = "True"
        arrValues(183) = "8"
        arrDone(184) = "True"
        arrValues(184) = "2"
        arrDone(185) = "False"
        arrValues(185) = "67A"
        arrDone(186) = "True"
        arrValues(186) = "D"
        arrDone(187) = "False"
        arrValues(187) = "467A"
        arrDone(188) = "False"
        arrValues(188) = "3ABC"
        arrDone(189) = "False"
        arrValues(189) = "346ABF"
        arrDone(190) = "True"
        arrValues(190) = "5"
        arrDone(191) = "True"
        arrValues(191) = "0"
        arrDone(192) = "False"
        arrValues(192) = "28ACF"
        arrDone(193) = "False"
        arrValues(193) = "28"
        arrDone(194) = "False"
        arrValues(194) = "2ACDF"
        arrDone(195) = "False"
        arrValues(195) = "ACD"
        arrDone(196) = "False"
        arrValues(196) = "178AB"
        arrDone(197) = "False"
        arrValues(197) = "2AD"
        arrDone(198) = "False"
        arrValues(198) = "28BD"
        arrDone(199) = "True"
        arrValues(199) = "3"
        arrDone(200) = "True"
        arrValues(200) = "4"
        arrDone(201) = "True"
        arrValues(201) = "5"
        arrDone(202) = "True"
        arrValues(202) = "9"
        arrDone(203) = "True"
        arrValues(203) = "E"
        arrDone(204) = "True"
        arrValues(204) = "6"
        arrDone(205) = "False"
        arrValues(205) = "0127ABDF"
        arrDone(206) = "False"
        arrValues(206) = "01ABCDF"
        arrDone(207) = "False"
        arrValues(207) = "128CF"
        arrDone(208) = "False"
        arrValues(208) = "2458AC"
        arrDone(209) = "False"
        arrValues(209) = "234568"
        arrDone(210) = "True"
        arrValues(210) = "1"
        arrDone(211) = "False"
        arrValues(211) = "3456ACD"
        arrDone(212) = "False"
        arrValues(212) = "578AB"
        arrDone(213) = "True"
        arrValues(213) = "0"
        arrDone(214) = "False"
        arrValues(214) = "2568BD"
        arrDone(215) = "False"
        arrValues(215) = "567A"
        arrDone(216) = "False"
        arrValues(216) = "67BD"
        arrDone(217) = "False"
        arrValues(217) = "367D"
        arrDone(218) = "False"
        arrValues(218) = "2367"
        arrDone(219) = "True"
        arrValues(219) = "F"
        arrDone(220) = "True"
        arrValues(220) = "9"
        arrDone(221) = "False"
        arrValues(221) = "23457ABD"
        arrDone(222) = "False"
        arrValues(222) = "34ABCDE"
        arrDone(223) = "False"
        arrValues(223) = "23458C"
        arrDone(224) = "True"
        arrValues(224) = "E"
        arrDone(225) = "False"
        arrValues(225) = "234568"
        arrDone(226) = "True"
        arrValues(226) = "B"
        arrDone(227) = "True"
        arrValues(227) = "0"
        arrDone(228) = "False"
        arrValues(228) = "578"
        arrDone(229) = "True"
        arrValues(229) = "C"
        arrDone(230) = "True"
        arrValues(230) = "F"
        arrDone(231) = "False"
        arrValues(231) = "5679"
        arrDone(232) = "True"
        arrValues(232) = "A"
        arrDone(233) = "True"
        arrValues(233) = "1"
        arrDone(234) = "False"
        arrValues(234) = "2367"
        arrDone(235) = "False"
        arrValues(235) = "67D"
        arrDone(236) = "False"
        arrValues(236) = "2357D"
        arrDone(237) = "False"
        arrValues(237) = "23457D"
        arrDone(238) = "False"
        arrValues(238) = "34D"
        arrDone(239) = "False"
        arrValues(239) = "23458"
        arrDone(240) = "True"
        arrValues(240) = "7"
        arrDone(241) = "False"
        arrValues(241) = "2356"
        arrDone(242) = "False"
        arrValues(242) = "2359ADF"
        arrDone(243) = "False"
        arrValues(243) = "3569AD"
        arrDone(244) = "True"
        arrValues(244) = "4"
        arrDone(245) = "False"
        arrValues(245) = "2569AD"
        arrDone(246) = "True"
        arrValues(246) = "E"
        arrDone(247) = "False"
        arrValues(247) = "1569A"
        arrDone(248) = "True"
        arrValues(248) = "8"
        arrDone(249) = "True"
        arrValues(249) = "0"
        arrDone(250) = "True"
        arrValues(250) = "C"
        arrDone(251) = "False"
        arrValues(251) = "6BD"
        arrDone(252) = "False"
        arrValues(252) = "1235ABD"
        arrDone(253) = "False"
        arrValues(253) = "1235ABDF"
        arrDone(254) = "False"
        arrValues(254) = "13ABDF"
        arrDone(255) = "False"
        arrValues(255) = "1235F"

        RefreshPuzzle()

    End Sub
    Private Sub LoadSolution3()

        arrDone(0) = "True"
        arrValues(0) = "2"
        arrDone(1) = "False"
        arrValues(1) = "58"
        arrDone(2) = "True"
        arrValues(2) = "6"
        arrDone(3) = "False"
        arrValues(3) = "0B"
        arrDone(4) = "False"
        arrValues(4) = "89"
        arrDone(5) = "True"
        arrValues(5) = "3"
        arrDone(6) = "True"
        arrValues(6) = "D"
        arrDone(7) = "False"
        arrValues(7) = "48AB"
        arrDone(8) = "True"
        arrValues(8) = "1"
        arrDone(9) = "False"
        arrValues(9) = "5789"
        arrDone(10) = "True"
        arrValues(10) = "C"
        arrDone(11) = "False"
        arrValues(11) = "078"
        arrDone(12) = "False"
        arrValues(12) = "045A"
        arrDone(13) = "False"
        arrValues(13) = "0459"
        arrDone(14) = "True"
        arrValues(14) = "F"
        arrDone(15) = "True"
        arrValues(15) = "E"
        arrDone(16) = "True"
        arrValues(16) = "D"
        arrDone(17) = "False"
        arrValues(17) = "58"
        arrDone(18) = "True"
        arrValues(18) = "9"
        arrDone(19) = "False"
        arrValues(19) = "013"
        arrDone(20) = "False"
        arrValues(20) = "128C"
        arrDone(21) = "True"
        arrValues(21) = "7"
        arrDone(22) = "False"
        arrValues(22) = "128C"
        arrDone(23) = "False"
        arrValues(23) = "8A"
        arrDone(24) = "False"
        arrValues(24) = "38A"
        arrDone(25) = "True"
        arrValues(25) = "F"
        arrDone(26) = "True"
        arrValues(26) = "4"
        arrDone(27) = "True"
        arrValues(27) = "E"
        arrDone(28) = "True"
        arrValues(28) = "6"
        arrDone(29) = "False"
        arrValues(29) = "05C"
        arrDone(30) = "True"
        arrValues(30) = "B"
        arrDone(31) = "False"
        arrValues(31) = "03A"
        arrDone(32) = "True"
        arrValues(32) = "E"
        arrDone(33) = "True"
        arrValues(33) = "A"
        arrDone(34) = "True"
        arrValues(34) = "C"
        arrDone(35) = "False"
        arrValues(35) = "03"
        arrDone(36) = "True"
        arrValues(36) = "6"
        arrDone(37) = "False"
        arrValues(37) = "49F"
        arrDone(38) = "False"
        arrValues(38) = "48F"
        arrDone(39) = "True"
        arrValues(39) = "5"
        arrDone(40) = "True"
        arrValues(40) = "2"
        arrDone(41) = "False"
        arrValues(41) = "89"
        arrDone(42) = "False"
        arrValues(42) = "03"
        arrDone(43) = "True"
        arrValues(43) = "B"
        arrDone(44) = "True"
        arrValues(44) = "D"
        arrDone(45) = "False"
        arrValues(45) = "049"
        arrDone(46) = "True"
        arrValues(46) = "1"
        arrDone(47) = "True"
        arrValues(47) = "7"
        arrDone(48) = "True"
        arrValues(48) = "F"
        arrDone(49) = "True"
        arrValues(49) = "4"
        arrDone(50) = "True"
        arrValues(50) = "7"
        arrDone(51) = "False"
        arrValues(51) = "13B"
        arrDone(52) = "False"
        arrValues(52) = "19C"
        arrDone(53) = "True"
        arrValues(53) = "0"
        arrDone(54) = "True"
        arrValues(54) = "E"
        arrDone(55) = "False"
        arrValues(55) = "AB"
        arrDone(56) = "False"
        arrValues(56) = "39A"
        arrDone(57) = "False"
        arrValues(57) = "359"
        arrDone(58) = "True"
        arrValues(58) = "6"
        arrDone(59) = "True"
        arrValues(59) = "D"
        arrDone(60) = "True"
        arrValues(60) = "2"
        arrDone(61) = "False"
        arrValues(61) = "59C"
        arrDone(62) = "True"
        arrValues(62) = "8"
        arrDone(63) = "False"
        arrValues(63) = "39A"
        arrDone(64) = "True"
        arrValues(64) = "8"
        arrDone(65) = "True"
        arrValues(65) = "0"
        arrDone(66) = "True"
        arrValues(66) = "2"
        arrDone(67) = "True"
        arrValues(67) = "D"
        arrDone(68) = "True"
        arrValues(68) = "3"
        arrDone(69) = "True"
        arrValues(69) = "5"
        arrDone(70) = "True"
        arrValues(70) = "A"
        arrDone(71) = "True"
        arrValues(71) = "F"
        arrDone(72) = "True"
        arrValues(72) = "B"
        arrDone(73) = "True"
        arrValues(73) = "C"
        arrDone(74) = "False"
        arrValues(74) = "1"
        arrDone(75) = "True"
        arrValues(75) = "4"
        arrDone(76) = "True"
        arrValues(76) = "E"
        arrDone(77) = "True"
        arrValues(77) = "6"
        arrDone(78) = "True"
        arrValues(78) = "7"
        arrDone(79) = "False"
        arrValues(79) = "19"
        arrDone(80) = "True"
        arrValues(80) = "B"
        arrDone(81) = "True"
        arrValues(81) = "E"
        arrDone(82) = "True"
        arrValues(82) = "5"
        arrDone(83) = "True"
        arrValues(83) = "F"
        arrDone(84) = "True"
        arrValues(84) = "0"
        arrDone(85) = "True"
        arrValues(85) = "C"
        arrDone(86) = "False"
        arrValues(86) = "14678"
        arrDone(87) = "False"
        arrValues(87) = "48"
        arrDone(88) = "False"
        arrValues(88) = "678"
        arrDone(89) = "False"
        arrValues(89) = "678D"
        arrDone(90) = "False"
        arrValues(90) = "17D"
        arrDone(91) = "True"
        arrValues(91) = "9"
        arrDone(92) = "False"
        arrValues(92) = "14"
        arrDone(93) = "True"
        arrValues(93) = "A"
        arrDone(94) = "True"
        arrValues(94) = "3"
        arrDone(95) = "True"
        arrValues(95) = "2"
        arrDone(96) = "True"
        arrValues(96) = "4"
        arrDone(97) = "True"
        arrValues(97) = "1"
        arrDone(98) = "True"
        arrValues(98) = "A"
        arrDone(99) = "True"
        arrValues(99) = "7"
        arrDone(100) = "True"
        arrValues(100) = "D"
        arrDone(101) = "True"
        arrValues(101) = "E"
        arrDone(102) = "True"
        arrValues(102) = "9"
        arrDone(103) = "True"
        arrValues(103) = "6"
        arrDone(104) = "True"
        arrValues(104) = "5"
        arrDone(105) = "False"
        arrValues(105) = "23"
        arrDone(106) = "True"
        arrValues(106) = "F"
        arrDone(107) = "False"
        arrValues(107) = "23"
        arrDone(108) = "True"
        arrValues(108) = "B"
        arrDone(109) = "True"
        arrValues(109) = "8"
        arrDone(110) = "True"
        arrValues(110) = "0"
        arrDone(111) = "True"
        arrValues(111) = "C"
        arrDone(112) = "True"
        arrValues(112) = "9"
        arrDone(113) = "True"
        arrValues(113) = "6"
        arrDone(114) = "True"
        arrValues(114) = "3"
        arrDone(115) = "True"
        arrValues(115) = "C"
        arrDone(116) = "True"
        arrValues(116) = "B"
        arrDone(117) = "False"
        arrValues(117) = "14"
        arrDone(118) = "False"
        arrValues(118) = "1478"
        arrDone(119) = "True"
        arrValues(119) = "2"
        arrDone(120) = "True"
        arrValues(120) = "E"
        arrDone(121) = "True"
        arrValues(121) = "A"
        arrDone(122) = "False"
        arrValues(122) = "017"
        arrDone(123) = "False"
        arrValues(123) = "0178"
        arrDone(124) = "False"
        arrValues(124) = "14"
        arrDone(125) = "True"
        arrValues(125) = "F"
        arrDone(126) = "True"
        arrValues(126) = "5"
        arrDone(127) = "True"
        arrValues(127) = "D"
        arrDone(128) = "True"
        arrValues(128) = "C"
        arrDone(129) = "True"
        arrValues(129) = "2"
        arrDone(130) = "True"
        arrValues(130) = "D"
        arrDone(131) = "True"
        arrValues(131) = "6"
        arrDone(132) = "False"
        arrValues(132) = "189F"
        arrDone(133) = "False"
        arrValues(133) = "149F"
        arrDone(134) = "True"
        arrValues(134) = "B"
        arrDone(135) = "True"
        arrValues(135) = "7"
        arrDone(136) = "True"
        arrValues(136) = "0"
        arrDone(137) = "False"
        arrValues(137) = "34589"
        arrDone(138) = "False"
        arrValues(138) = "135"
        arrDone(139) = "False"
        arrValues(139) = "138"
        arrDone(140) = "False"
        arrValues(140) = "145AF"
        arrDone(141) = "True"
        arrValues(141) = "E"
        arrDone(142) = "False"
        arrValues(142) = "49"
        arrDone(143) = "False"
        arrValues(143) = "189AF"
        arrDone(144) = "True"
        arrValues(144) = "1"
        arrDone(145) = "True"
        arrValues(145) = "F"
        arrDone(146) = "True"
        arrValues(146) = "E"
        arrDone(147) = "True"
        arrValues(147) = "9"
        arrDone(148) = "False"
        arrValues(148) = "28"
        arrDone(149) = "True"
        arrValues(149) = "A"
        arrDone(150) = "False"
        arrValues(150) = "0248"
        arrDone(151) = "True"
        arrValues(151) = "D"
        arrDone(152) = "False"
        arrValues(152) = "478C"
        arrDone(153) = "False"
        arrValues(153) = "24578"
        arrDone(154) = "False"
        arrValues(154) = "257"
        arrDone(155) = "False"
        arrValues(155) = "278C"
        arrDone(156) = "True"
        arrValues(156) = "3"
        arrDone(157) = "False"
        arrValues(157) = "0245"
        arrDone(158) = "True"
        arrValues(158) = "6"
        arrDone(159) = "True"
        arrValues(159) = "B"
        arrDone(160) = "True"
        arrValues(160) = "A"
        arrDone(161) = "True"
        arrValues(161) = "B"
        arrDone(162) = "True"
        arrValues(162) = "0"
        arrDone(163) = "True"
        arrValues(163) = "5"
        arrDone(164) = "True"
        arrValues(164) = "E"
        arrDone(165) = "False"
        arrValues(165) = "1249F"
        arrDone(166) = "False"
        arrValues(166) = "1248F"
        arrDone(167) = "True"
        arrValues(167) = "3"
        arrDone(168) = "False"
        arrValues(168) = "489"
        arrDone(169) = "False"
        arrValues(169) = "2489D"
        arrDone(170) = "False"
        arrValues(170) = "12D"
        arrDone(171) = "True"
        arrValues(171) = "6"
        arrDone(172) = "True"
        arrValues(172) = "C"
        arrDone(173) = "True"
        arrValues(173) = "7"
        arrDone(174) = "False"
        arrValues(174) = "49"
        arrDone(175) = "False"
        arrValues(175) = "189F"
        arrDone(176) = "True"
        arrValues(176) = "7"
        arrDone(177) = "True"
        arrValues(177) = "3"
        arrDone(178) = "True"
        arrValues(178) = "4"
        arrDone(179) = "True"
        arrValues(179) = "8"
        arrDone(180) = "True"
        arrValues(180) = "5"
        arrDone(181) = "False"
        arrValues(181) = "1269"
        arrDone(182) = "False"
        arrValues(182) = "0126"
        arrDone(183) = "True"
        arrValues(183) = "C"
        arrDone(184) = "True"
        arrValues(184) = "F"
        arrDone(185) = "True"
        arrValues(185) = "B"
        arrDone(186) = "True"
        arrValues(186) = "E"
        arrDone(187) = "True"
        arrValues(187) = "A"
        arrDone(188) = "False"
        arrValues(188) = "01"
        arrDone(189) = "False"
        arrValues(189) = "029"
        arrDone(190) = "True"
        arrValues(190) = "D"
        arrDone(191) = "False"
        arrValues(191) = "019"
        arrDone(192) = "True"
        arrValues(192) = "5"
        arrDone(193) = "True"
        arrValues(193) = "9"
        arrDone(194) = "True"
        arrValues(194) = "1"
        arrDone(195) = "True"
        arrValues(195) = "4"
        arrDone(196) = "False"
        arrValues(196) = "27CF"
        arrDone(197) = "True"
        arrValues(197) = "B"
        arrDone(198) = "False"
        arrValues(198) = "237CF"
        arrDone(199) = "True"
        arrValues(199) = "E"
        arrDone(200) = "False"
        arrValues(200) = "37C"
        arrDone(201) = "True"
        arrValues(201) = "0"
        arrDone(202) = "True"
        arrValues(202) = "8"
        arrDone(203) = "False"
        arrValues(203) = "237CF"
        arrDone(204) = "False"
        arrValues(204) = "7F"
        arrDone(205) = "True"
        arrValues(205) = "D"
        arrDone(206) = "True"
        arrValues(206) = "A"
        arrDone(207) = "True"
        arrValues(207) = "6"
        arrDone(208) = "True"
        arrValues(208) = "0"
        arrDone(209) = "True"
        arrValues(209) = "7"
        arrDone(210) = "True"
        arrValues(210) = "B"
        arrDone(211) = "True"
        arrValues(211) = "2"
        arrDone(212) = "True"
        arrValues(212) = "4"
        arrDone(213) = "False"
        arrValues(213) = "6F"
        arrDone(214) = "False"
        arrValues(214) = "36F"
        arrDone(215) = "True"
        arrValues(215) = "9"
        arrDone(216) = "True"
        arrValues(216) = "D"
        arrDone(217) = "True"
        arrValues(217) = "E"
        arrDone(218) = "True"
        arrValues(218) = "A"
        arrDone(219) = "False"
        arrValues(219) = "3F"
        arrDone(220) = "True"
        arrValues(220) = "8"
        arrDone(221) = "True"
        arrValues(221) = "1"
        arrDone(222) = "True"
        arrValues(222) = "C"
        arrDone(223) = "True"
        arrValues(223) = "5"
        arrDone(224) = "True"
        arrValues(224) = "3"
        arrDone(225) = "True"
        arrValues(225) = "C"
        arrDone(226) = "True"
        arrValues(226) = "8"
        arrDone(227) = "True"
        arrValues(227) = "E"
        arrDone(228) = "True"
        arrValues(228) = "A"
        arrDone(229) = "True"
        arrValues(229) = "D"
        arrDone(230) = "True"
        arrValues(230) = "5"
        arrDone(231) = "True"
        arrValues(231) = "1"
        arrDone(232) = "False"
        arrValues(232) = "467"
        arrDone(233) = "False"
        arrValues(233) = "467"
        arrDone(234) = "True"
        arrValues(234) = "9"
        arrDone(235) = "False"
        arrValues(235) = "7F"
        arrDone(236) = "False"
        arrValues(236) = "07F"
        arrDone(237) = "True"
        arrValues(237) = "B"
        arrDone(238) = "True"
        arrValues(238) = "2"
        arrDone(239) = "False"
        arrValues(239) = "0F"
        arrDone(240) = "True"
        arrValues(240) = "6"
        arrDone(241) = "True"
        arrValues(241) = "D"
        arrDone(242) = "True"
        arrValues(242) = "F"
        arrDone(243) = "True"
        arrValues(243) = "A"
        arrDone(244) = "False"
        arrValues(244) = "27C"
        arrDone(245) = "True"
        arrValues(245) = "8"
        arrDone(246) = "False"
        arrValues(246) = "27C"
        arrDone(247) = "True"
        arrValues(247) = "0"
        arrDone(248) = "False"
        arrValues(248) = "7C"
        arrDone(249) = "True"
        arrValues(249) = "1"
        arrDone(250) = "True"
        arrValues(250) = "B"
        arrDone(251) = "True"
        arrValues(251) = "5"
        arrDone(252) = "True"
        arrValues(252) = "9"
        arrDone(253) = "True"
        arrValues(253) = "3"
        arrDone(254) = "True"
        arrValues(254) = "E"
        arrDone(255) = "True"
        arrValues(255) = "4"

        RefreshPuzzle()
    End Sub

    Private Sub ButtonLoadSolution4_Click_1(sender As Object, e As EventArgs) Handles ButtonLoadSolution4.Click
        'A five star puzzle

        arrDone(0) = "False"
        arrValues(0) = "02468AC"
        arrDone(1) = "False"
        arrValues(1) = "02468AC"
        arrDone(2) = "False"
        arrValues(2) = "02468AF"
        arrDone(3) = "False"
        arrValues(3) = "48AC"
        arrDone(4) = "True"
        arrValues(4) = "3"
        arrDone(5) = "True"
        arrValues(5) = "D"
        arrDone(6) = "True"
        arrValues(6) = "9"
        arrDone(7) = "True"
        arrValues(7) = "B"
        arrDone(8) = "True"
        arrValues(8) = "1"
        arrDone(9) = "False"
        arrValues(9) = "48AC"
        arrDone(10) = "False"
        arrValues(10) = "4AC"
        arrDone(11) = "False"
        arrValues(11) = "04C"
        arrDone(12) = "True"
        arrValues(12) = "E"
        arrDone(13) = "False"
        arrValues(13) = "67AC"
        arrDone(14) = "False"
        arrValues(14) = "4567C"
        arrDone(15) = "False"
        arrValues(15) = "067AC"
        arrDone(16) = "False"
        arrValues(16) = "01468CE"
        arrDone(17) = "True"
        arrValues(17) = "D"
        arrDone(18) = "False"
        arrValues(18) = "0468EF"
        arrDone(19) = "False"
        arrValues(19) = "348CE"
        arrDone(20) = "False"
        arrValues(20) = "468EF"
        arrDone(21) = "True"
        arrValues(21) = "A"
        arrDone(22) = "False"
        arrValues(22) = "56EF"
        arrDone(23) = "False"
        arrValues(23) = "045EF"
        arrDone(24) = "False"
        arrValues(24) = "0389C"
        arrDone(25) = "True"
        arrValues(25) = "7"
        arrDone(26) = "True"
        arrValues(26) = "2"
        arrDone(27) = "False"
        arrValues(27) = "04C"
        arrDone(28) = "True"
        arrValues(28) = "B"
        arrDone(29) = "False"
        arrValues(29) = "369C"
        arrDone(30) = "False"
        arrValues(30) = "34569C"
        arrDone(31) = "False"
        arrValues(31) = "069C"
        arrDone(32) = "True"
        arrValues(32) = "9"
        arrDone(33) = "False"
        arrValues(33) = "04ABE"
        arrDone(34) = "False"
        arrValues(34) = "04AE"
        arrDone(35) = "False"
        arrValues(35) = "34ABE"
        arrDone(36) = "True"
        arrValues(36) = "C"
        arrDone(37) = "True"
        arrValues(37) = "1"
        arrDone(38) = "False"
        arrValues(38) = "E"
        arrDone(39) = "True"
        arrValues(39) = "7"
        arrDone(40) = "True"
        arrValues(40) = "F"
        arrDone(41) = "True"
        arrValues(41) = "6"
        arrDone(42) = "False"
        arrValues(42) = "34AB"
        arrDone(43) = "True"
        arrValues(43) = "5"
        arrDone(44) = "True"
        arrValues(44) = "2"
        arrDone(45) = "True"
        arrValues(45) = "8"
        arrDone(46) = "False"
        arrValues(46) = "34D"
        arrDone(47) = "False"
        arrValues(47) = "0AD"
        arrDone(48) = "True"
        arrValues(48) = "7"
        arrDone(49) = "False"
        arrValues(49) = "02468ABC"
        arrDone(50) = "True"
        arrValues(50) = "5"
        arrDone(51) = "False"
        arrValues(51) = "348ABC"
        arrDone(52) = "False"
        arrValues(52) = "468"
        arrDone(53) = "False"
        arrValues(53) = "0246"
        arrDone(54) = "False"
        arrValues(54) = "6"
        arrDone(55) = "False"
        arrValues(55) = "024"
        arrDone(56) = "False"
        arrValues(56) = "0389ABC"
        arrDone(57) = "False"
        arrValues(57) = "489ACD"
        arrDone(58) = "True"
        arrValues(58) = "E"
        arrDone(59) = "False"
        arrValues(59) = "04BCD"
        arrDone(60) = "True"
        arrValues(60) = "F"
        arrDone(61) = "True"
        arrValues(61) = "1"
        arrDone(62) = "False"
        arrValues(62) = "3469CD"
        arrDone(63) = "False"
        arrValues(63) = "069ACD"
        arrDone(64) = "False"
        arrValues(64) = "124BCE"
        arrDone(65) = "False"
        arrValues(65) = "124BCE"
        arrDone(66) = "False"
        arrValues(66) = "249E"
        arrDone(67) = "True"
        arrValues(67) = "0"
        arrDone(68) = "True"
        arrValues(68) = "D"
        arrDone(69) = "False"
        arrValues(69) = "479E"
        arrDone(70) = "False"
        arrValues(70) = "1BE"
        arrDone(71) = "True"
        arrValues(71) = "6"
        arrDone(72) = "False"
        arrValues(72) = "7BC"
        arrDone(73) = "False"
        arrValues(73) = "24CE"
        arrDone(74) = "True"
        arrValues(74) = "8"
        arrDone(75) = "False"
        arrValues(75) = "47BCE"
        arrDone(76) = "False"
        arrValues(76) = "379C"
        arrDone(77) = "True"
        arrValues(77) = "5"
        arrDone(78) = "True"
        arrValues(78) = "A"
        arrDone(79) = "True"
        arrValues(79) = "F"
        arrDone(80) = "False"
        arrValues(80) = "1468ABCDE"
        arrDone(81) = "True"
        arrValues(81) = "3"
        arrDone(82) = "False"
        arrValues(82) = "4689ADE"
        arrDone(83) = "False"
        arrValues(83) = "48ABCDE"
        arrDone(84) = "False"
        arrValues(84) = "1479BEF"
        arrDone(85) = "False"
        arrValues(85) = "479E"
        arrDone(86) = "False"
        arrValues(86) = "15BEF"
        arrDone(87) = "False"
        arrValues(87) = "1459EF"
        arrDone(88) = "False"
        arrValues(88) = "567ABC"
        arrDone(89) = "False"
        arrValues(89) = "4ACDE"
        arrDone(90) = "True"
        arrValues(90) = "0"
        arrDone(91) = "False"
        arrValues(91) = "467BCDE"
        arrDone(92) = "False"
        arrValues(92) = "789CD"
        arrDone(93) = "True"
        arrValues(93) = "2"
        arrDone(94) = "False"
        arrValues(94) = "179BCD"
        arrDone(95) = "False"
        arrValues(95) = "1789BCD"
        arrDone(96) = "False"
        arrValues(96) = "18ABD"
        arrDone(97) = "False"
        arrValues(97) = "18AB"
        arrDone(98) = "False"
        arrValues(98) = "8AD"
        arrDone(99) = "True"
        arrValues(99) = "7"
        arrDone(100) = "True"
        arrValues(100) = "2"
        arrDone(101) = "True"
        arrValues(101) = "3"
        arrDone(102) = "True"
        arrValues(102) = "0"
        arrDone(103) = "True"
        arrValues(103) = "C"
        arrDone(104) = "False"
        arrValues(104) = "5AB"
        arrDone(105) = "True"
        arrValues(105) = "F"
        arrDone(106) = "False"
        arrValues(106) = "AB"
        arrDone(107) = "True"
        arrValues(107) = "9"
        arrDone(108) = "True"
        arrValues(108) = "6"
        arrDone(109) = "True"
        arrValues(109) = "4"
        arrDone(110) = "False"
        arrValues(110) = "1BD"
        arrDone(111) = "True"
        arrValues(111) = "E"
        arrDone(112) = "True"
        arrValues(112) = "F"
        arrDone(113) = "False"
        arrValues(113) = "246BCE"
        arrDone(114) = "False"
        arrValues(114) = "2469DE"
        arrDone(115) = "True"
        arrValues(115) = "5"
        arrDone(116) = "True"
        arrValues(116) = "A"
        arrDone(117) = "False"
        arrValues(117) = "479E"
        arrDone(118) = "True"
        arrValues(118) = "8"
        arrDone(119) = "False"
        arrValues(119) = "49E"
        arrDone(120) = "False"
        arrValues(120) = "67BC"
        arrDone(121) = "True"
        arrValues(121) = "3"
        arrDone(122) = "True"
        arrValues(122) = "1"
        arrDone(123) = "False"
        arrValues(123) = "467BCDE"
        arrDone(124) = "False"
        arrValues(124) = "79CD"
        arrDone(125) = "False"
        arrValues(125) = "79BCD"
        arrDone(126) = "True"
        arrValues(126) = "0"
        arrDone(127) = "False"
        arrValues(127) = "79BCD"
        arrDone(128) = "False"
        arrValues(128) = "068AE"
        arrDone(129) = "True"
        arrValues(129) = "F"
        arrDone(130) = "False"
        arrValues(130) = "068AE"
        arrDone(131) = "False"
        arrValues(131) = "8AE"
        arrDone(132) = "False"
        arrValues(132) = "679E"
        arrDone(133) = "True"
        arrValues(133) = "B"
        arrDone(134) = "True"
        arrValues(134) = "D"
        arrDone(135) = "False"
        arrValues(135) = "039E"
        arrDone(136) = "False"
        arrValues(136) = "0367C"
        arrDone(137) = "True"
        arrValues(137) = "5"
        arrDone(138) = "False"
        arrValues(138) = "37C"
        arrDone(139) = "True"
        arrValues(139) = "2"
        arrDone(140) = "True"
        arrValues(140) = "1"
        arrDone(141) = "False"
        arrValues(141) = "3679CE"
        arrDone(142) = "False"
        arrValues(142) = "3679CE"
        arrDone(143) = "True"
        arrValues(143) = "4"
        arrDone(144) = "True"
        arrValues(144) = "3"
        arrDone(145) = "False"
        arrValues(145) = "06E"
        arrDone(146) = "True"
        arrValues(146) = "7"
        arrDone(147) = "True"
        arrValues(147) = "1"
        arrDone(148) = "True"
        arrValues(148) = "5"
        arrDone(149) = "False"
        arrValues(149) = "06E"
        arrDone(150) = "True"
        arrValues(150) = "2"
        arrDone(151) = "False"
        arrValues(151) = "0EF"
        arrDone(152) = "True"
        arrValues(152) = "4"
        arrDone(153) = "True"
        arrValues(153) = "B"
        arrDone(154) = "True"
        arrValues(154) = "9"
        arrDone(155) = "True"
        arrValues(155) = "8"
        arrDone(156) = "True"
        arrValues(156) = "A"
        arrDone(157) = "False"
        arrValues(157) = "6CDE"
        arrDone(158) = "False"
        arrValues(158) = "6CDE"
        arrDone(159) = "False"
        arrValues(159) = "6CD"
        arrDone(160) = "False"
        arrValues(160) = "02468ADE"
        arrDone(161) = "False"
        arrValues(161) = "02468AE"
        arrDone(162) = "True"
        arrValues(162) = "B"
        arrDone(163) = "False"
        arrValues(163) = "48ADE"
        arrDone(164) = "False"
        arrValues(164) = "14679E"
        arrDone(165) = "True"
        arrValues(165) = "C"
        arrDone(166) = "False"
        arrValues(166) = "136AE"
        arrDone(167) = "False"
        arrValues(167) = "01349E"
        arrDone(168) = "False"
        arrValues(168) = "0367"
        arrDone(169) = "False"
        arrValues(169) = "1E"
        arrDone(170) = "False"
        arrValues(170) = "37"
        arrDone(171) = "False"
        arrValues(171) = "0167E"
        arrDone(172) = "False"
        arrValues(172) = "35789D"
        arrDone(173) = "False"
        arrValues(173) = "3679DE"
        arrDone(174) = "True"
        arrValues(174) = "F"
        arrDone(175) = "False"
        arrValues(175) = "26789D"
        arrDone(176) = "True"
        arrValues(176) = "5"
        arrDone(177) = "True"
        arrValues(177) = "9"
        arrDone(178) = "True"
        arrValues(178) = "C"
        arrDone(179) = "False"
        arrValues(179) = "4E"
        arrDone(180) = "False"
        arrValues(180) = "1467EF"
        arrDone(181) = "True"
        arrValues(181) = "8"
        arrDone(182) = "False"
        arrValues(182) = "136EF"
        arrDone(183) = "False"
        arrValues(183) = "134EF"
        arrDone(184) = "True"
        arrValues(184) = "D"
        arrDone(185) = "False"
        arrValues(185) = "1E"
        arrDone(186) = "False"
        arrValues(186) = "37F"
        arrDone(187) = "True"
        arrValues(187) = "A"
        arrDone(188) = "True"
        arrValues(188) = "0"
        arrDone(189) = "False"
        arrValues(189) = "367BE"
        arrDone(190) = "False"
        arrValues(190) = "2367BE"
        arrDone(191) = "False"
        arrValues(191) = "267B"
        arrDone(192) = "False"
        arrValues(192) = "48ABCDE"
        arrDone(193) = "False"
        arrValues(193) = "478ABCE"
        arrDone(194) = "True"
        arrValues(194) = "1"
        arrDone(195) = "True"
        arrValues(195) = "F"
        arrDone(196) = "False"
        arrValues(196) = "69BE"
        arrDone(197) = "True"
        arrValues(197) = "5"
        arrDone(198) = "False"
        arrValues(198) = "6BCE"
        arrDone(199) = "False"
        arrValues(199) = "29DE"
        arrDone(200) = "False"
        arrValues(200) = "789ABC"
        arrDone(201) = "False"
        arrValues(201) = "489AC"
        arrDone(202) = "False"
        arrValues(202) = "47ABC"
        arrDone(203) = "False"
        arrValues(203) = "47BC"
        arrDone(204) = "False"
        arrValues(204) = "479CD"
        arrDone(205) = "True"
        arrValues(205) = "0"
        arrDone(206) = "False"
        arrValues(206) = "24679BCDE"
        arrDone(207) = "True"
        arrValues(207) = "3"
        arrDone(208) = "False"
        arrValues(208) = "8BC"
        arrDone(209) = "False"
        arrValues(209) = "78BC"
        arrDone(210) = "True"
        arrValues(210) = "3"
        arrDone(211) = "True"
        arrValues(211) = "6"
        arrDone(212) = "True"
        arrValues(212) = "0"
        arrDone(213) = "False"
        arrValues(213) = "29"
        arrDone(214) = "True"
        arrValues(214) = "4"
        arrDone(215) = "True"
        arrValues(215) = "A"
        arrDone(216) = "True"
        arrValues(216) = "E"
        arrDone(217) = "False"
        arrValues(217) = "189C"
        arrDone(218) = "True"
        arrValues(218) = "D"
        arrDone(219) = "True"
        arrValues(219) = "F"
        arrDone(220) = "False"
        arrValues(220) = "79C"
        arrDone(221) = "False"
        arrValues(221) = "79BC"
        arrDone(222) = "False"
        arrValues(222) = "1279BC"
        arrDone(223) = "True"
        arrValues(223) = "5"
        arrDone(224) = "False"
        arrValues(224) = "04ABCDE"
        arrDone(225) = "False"
        arrValues(225) = "045ABCE"
        arrDone(226) = "False"
        arrValues(226) = "04ADE"
        arrDone(227) = "True"
        arrValues(227) = "2"
        arrDone(228) = "False"
        arrValues(228) = "19BE"
        arrDone(229) = "True"
        arrValues(229) = "F"
        arrDone(230) = "True"
        arrValues(230) = "7"
        arrDone(231) = "False"
        arrValues(231) = "139DE"
        arrDone(232) = "False"
        arrValues(232) = "9ABC"
        arrDone(233) = "False"
        arrValues(233) = "149AC"
        arrDone(234) = "True"
        arrValues(234) = "6"
        arrDone(235) = "False"
        arrValues(235) = "14BC"
        arrDone(236) = "False"
        arrValues(236) = "49CD"
        arrDone(237) = "False"
        arrValues(237) = "9ABCDE"
        arrDone(238) = "True"
        arrValues(238) = "8"
        arrDone(239) = "False"
        arrValues(239) = "19ABCD"
        arrDone(240) = "False"
        arrValues(240) = "4ABCDE"
        arrDone(241) = "False"
        arrValues(241) = "47ABCE"
        arrDone(242) = "False"
        arrValues(242) = "4ADE"
        arrDone(243) = "True"
        arrValues(243) = "9"
        arrDone(244) = "False"
        arrValues(244) = "16BE"
        arrDone(245) = "False"
        arrValues(245) = "6E"
        arrDone(246) = "False"
        arrValues(246) = "16BCE"
        arrDone(247) = "True"
        arrValues(247) = "8"
        arrDone(248) = "True"
        arrValues(248) = "2"
        arrDone(249) = "True"
        arrValues(249) = "0"
        arrDone(250) = "True"
        arrValues(250) = "5"
        arrDone(251) = "True"
        arrValues(251) = "3"
        arrDone(252) = "False"
        arrValues(252) = "47CD"
        arrDone(253) = "False"
        arrValues(253) = "67ABCDEF"
        arrDone(254) = "False"
        arrValues(254) = "1467BCDE"
        arrDone(255) = "False"
        arrValues(255) = "167ABCD"

        'TextBox3.Text = "Sq3 quad 048C -> R1C10 = 8; R1 duple 4C, R1C12 = 0; remove others; C7 tuple 1BC, remove others; Sq6 tuple 479, R8C8 = E; remove others;"
        Hints = "Sq3 quad 048C -> R1C10 = 8; R1 duple 4C, R1C12 = 0; remove others; C7 tuple 1BC, remove others; Sq6 tuple 479, R8C8 = E; remove others;Sq11 duple 1E ->;Sq12 has duples CD and 6B ->; Sq9 R9 has tuple 68A ->;R2 has Quad 0369 -> R2C16 = 0; auto; R6 has duple 79 ->; auto remainder"
        RefreshPuzzle()

    End Sub

    Private Sub LoadSolution4A()

        arrDone(0) = "False"
        arrValues(0) = "368BCDF"
        arrDone(1) = "False"
        arrValues(1) = "368BD"
        arrDone(2) = "False"
        arrValues(2) = "68BCD"
        arrDone(3) = "False"
        arrValues(3) = "68B"
        arrDone(4) = "False"
        arrValues(4) = "256B"
        arrDone(5) = "False"
        arrValues(5) = "35BCE"
        arrDone(6) = "False"
        arrValues(6) = "26BDE"
        arrDone(7) = "False"
        arrValues(7) = "5CE"
        arrDone(8) = "True"
        arrValues(8) = "0"
        arrDone(9) = "True"
        arrValues(9) = "1"
        arrDone(10) = "False"
        arrValues(10) = "268BDEF"
        arrDone(11) = "True"
        arrValues(11) = "9"
        arrDone(12) = "True"
        arrValues(12) = "A"
        arrDone(13) = "True"
        arrValues(13) = "4"
        arrDone(14) = "False"
        arrValues(14) = "368D"
        arrDone(15) = "True"
        arrValues(15) = "7"
        arrDone(16) = "True"
        arrValues(16) = "0"
        arrDone(17) = "True"
        arrValues(17) = "2"
        arrDone(18) = "False"
        arrValues(18) = "46789D"
        arrDone(19) = "False"
        arrValues(19) = "6789"
        arrDone(20) = "True"
        arrValues(20) = "A"
        arrDone(21) = "False"
        arrValues(21) = "134E"
        arrDone(22) = "True"
        arrValues(22) = "F"
        arrDone(23) = "False"
        arrValues(23) = "1479E"
        arrDone(24) = "False"
        arrValues(24) = "68DE"
        arrDone(25) = "False"
        arrValues(25) = "68E"
        arrDone(26) = "False"
        arrValues(26) = "68DE"
        arrDone(27) = "True"
        arrValues(27) = "C"
        arrDone(28) = "True"
        arrValues(28) = "5"
        arrDone(29) = "False"
        arrValues(29) = "3689D"
        arrDone(30) = "True"
        arrValues(30) = "B"
        arrDone(31) = "False"
        arrValues(31) = "3689D"
        arrDone(32) = "True"
        arrValues(32) = "1"
        arrDone(33) = "True"
        arrValues(33) = "A"
        arrDone(34) = "False"
        arrValues(34) = "4679BCD"
        arrDone(35) = "True"
        arrValues(35) = "5"
        arrDone(36) = "False"
        arrValues(36) = "2469B"
        arrDone(37) = "False"
        arrValues(37) = "04BC"
        arrDone(38) = "True"
        arrValues(38) = "8"
        arrDone(39) = "False"
        arrValues(39) = "0479C"
        arrDone(40) = "False"
        arrValues(40) = "6BD"
        arrDone(41) = "True"
        arrValues(41) = "3"
        arrDone(42) = "False"
        arrValues(42) = "26BD"
        arrDone(43) = "False"
        arrValues(43) = "26D"
        arrDone(44) = "True"
        arrValues(44) = "F"
        arrDone(45) = "True"
        arrValues(45) = "E"
        arrDone(46) = "False"
        arrValues(46) = "069D"
        arrDone(47) = "False"
        arrValues(47) = "0269CD"
        arrDone(48) = "False"
        arrValues(48) = "368BCDF"
        arrDone(49) = "True"
        arrValues(49) = "E"
        arrDone(50) = "False"
        arrValues(50) = "689BCD"
        arrDone(51) = "False"
        arrValues(51) = "689B"
        arrDone(52) = "False"
        arrValues(52) = "269B"
        arrDone(53) = "False"
        arrValues(53) = "03BC"
        arrDone(54) = "False"
        arrValues(54) = "0269BD"
        arrDone(55) = "False"
        arrValues(55) = "09C"
        arrDone(56) = "False"
        arrValues(56) = "68ABDF"
        arrDone(57) = "True"
        arrValues(57) = "4"
        arrDone(58) = "True"
        arrValues(58) = "5"
        arrDone(59) = "True"
        arrValues(59) = "7"
        arrDone(60) = "False"
        arrValues(60) = "2689CD"
        arrDone(61) = "False"
        arrValues(61) = "023689CD"
        arrDone(62) = "True"
        arrValues(62) = "1"
        arrDone(63) = "False"
        arrValues(63) = "023689CD"
        arrDone(64) = "False"
        arrValues(64) = "678ABD"
        arrDone(65) = "True"
        arrValues(65) = "C"
        arrDone(66) = "False"
        arrValues(66) = "678ABD"
        arrDone(67) = "False"
        arrValues(67) = "678AB"
        arrDone(68) = "False"
        arrValues(68) = "4F"
        arrDone(69) = "False"
        arrValues(69) = "4AEF"
        arrDone(70) = "True"
        arrValues(70) = "1"
        arrDone(71) = "False"
        arrValues(71) = "47AEF"
        arrDone(72) = "True"
        arrValues(72) = "2"
        arrDone(73) = "True"
        arrValues(73) = "5"
        arrDone(74) = "True"
        arrValues(74) = "0"
        arrDone(75) = "False"
        arrValues(75) = "6DE"
        arrDone(76) = "True"
        arrValues(76) = "3"
        arrDone(77) = "False"
        arrValues(77) = "6789AD"
        arrDone(78) = "False"
        arrValues(78) = "689D"
        arrDone(79) = "False"
        arrValues(79) = "4689D"
        arrDone(80) = "False"
        arrValues(80) = "2567AD"
        arrDone(81) = "False"
        arrValues(81) = "56D"
        arrDone(82) = "False"
        arrValues(82) = "267AD"
        arrDone(83) = "True"
        arrValues(83) = "0"
        arrDone(84) = "True"
        arrValues(84) = "8"
        arrDone(85) = "True"
        arrValues(85) = "9"
        arrDone(86) = "False"
        arrValues(86) = "27"
        arrDone(87) = "True"
        arrValues(87) = "B"
        arrDone(88) = "True"
        arrValues(88) = "3"
        arrDone(89) = "False"
        arrValues(89) = "6AF"
        arrDone(90) = "False"
        arrValues(90) = "6ACDF"
        arrDone(91) = "False"
        arrValues(91) = "6D"
        arrDone(92) = "True"
        arrValues(92) = "1"
        arrDone(93) = "False"
        arrValues(93) = "2567ACD"
        arrDone(94) = "True"
        arrValues(94) = "E"
        arrDone(95) = "False"
        arrValues(95) = "246CD"
        arrDone(96) = "True"
        arrValues(96) = "E"
        arrDone(97) = "False"
        arrValues(97) = "368"
        arrDone(98) = "False"
        arrValues(98) = "12678A"
        arrDone(99) = "False"
        arrValues(99) = "12678A"
        arrDone(100) = "True"
        arrValues(100) = "0"
        arrDone(101) = "True"
        arrValues(101) = "D"
        arrDone(102) = "True"
        arrValues(102) = "5"
        arrDone(103) = "False"
        arrValues(103) = "7AC"
        arrDone(104) = "True"
        arrValues(104) = "4"
        arrDone(105) = "False"
        arrValues(105) = "69A"
        arrDone(106) = "False"
        arrValues(106) = "69AC"
        arrDone(107) = "False"
        arrValues(107) = "16"
        arrDone(108) = "False"
        arrValues(108) = "26789C"
        arrDone(109) = "True"
        arrValues(109) = "B"
        arrDone(110) = "True"
        arrValues(110) = "F"
        arrDone(111) = "False"
        arrValues(111) = "2689C"
        arrDone(112) = "True"
        arrValues(112) = "4"
        arrDone(113) = "True"
        arrValues(113) = "9"
        arrDone(114) = "True"
        arrValues(114) = "F"
        arrDone(115) = "False"
        arrValues(115) = "127AB"
        arrDone(116) = "True"
        arrValues(116) = "3"
        arrDone(117) = "True"
        arrValues(117) = "6"
        arrDone(118) = "False"
        arrValues(118) = "27E"
        arrDone(119) = "False"
        arrValues(119) = "7ACE"
        arrDone(120) = "False"
        arrValues(120) = "17ABCDE"
        arrDone(121) = "False"
        arrValues(121) = "AE"
        arrDone(122) = "False"
        arrValues(122) = "ABCDE"
        arrDone(123) = "True"
        arrValues(123) = "8"
        arrDone(124) = "False"
        arrValues(124) = "27CD"
        arrDone(125) = "False"
        arrValues(125) = "0257ACD"
        arrDone(126) = "False"
        arrValues(126) = "05D"
        arrDone(127) = "False"
        arrValues(127) = "02CD"
        arrDone(128) = "False"
        arrValues(128) = "268BD"
        arrDone(129) = "False"
        arrValues(129) = "068BD"
        arrDone(130) = "False"
        arrValues(130) = "012689BD"
        arrDone(131) = "False"
        arrValues(131) = "12689BE"
        arrDone(132) = "True"
        arrValues(132) = "C"
        arrDone(133) = "False"
        arrValues(133) = "0148BE"
        arrDone(134) = "False"
        arrValues(134) = "069BE"
        arrDone(135) = "False"
        arrValues(135) = "01489E"
        arrDone(136) = "False"
        arrValues(136) = "689E"
        arrDone(137) = "False"
        arrValues(137) = "02689E"
        arrDone(138) = "True"
        arrValues(138) = "3"
        arrDone(139) = "True"
        arrValues(139) = "A"
        arrDone(140) = "False"
        arrValues(140) = "24689BD"
        arrDone(141) = "True"
        arrValues(141) = "F"
        arrDone(142) = "True"
        arrValues(142) = "7"
        arrDone(143) = "True"
        arrValues(143) = "5"
        arrDone(144) = "False"
        arrValues(144) = "68BD"
        arrDone(145) = "True"
        arrValues(145) = "4"
        arrDone(146) = "True"
        arrValues(146) = "5"
        arrDone(147) = "False"
        arrValues(147) = "1689BE"
        arrDone(148) = "False"
        arrValues(148) = "169B"
        arrDone(149) = "False"
        arrValues(149) = "018BE"
        arrDone(150) = "False"
        arrValues(150) = "069BE"
        arrDone(151) = "True"
        arrValues(151) = "2"
        arrDone(152) = "False"
        arrValues(152) = "689E"
        arrDone(153) = "True"
        arrValues(153) = "C"
        arrDone(154) = "True"
        arrValues(154) = "7"
        arrDone(155) = "True"
        arrValues(155) = "F"
        arrDone(156) = "False"
        arrValues(156) = "689BD"
        arrDone(157) = "False"
        arrValues(157) = "013689D"
        arrDone(158) = "False"
        arrValues(158) = "03689D"
        arrDone(159) = "True"
        arrValues(159) = "A"
        arrDone(160) = "False"
        arrValues(160) = "268AC"
        arrDone(161) = "True"
        arrValues(161) = "7"
        arrDone(162) = "False"
        arrValues(162) = "02689AC"
        arrDone(163) = "True"
        arrValues(163) = "3"
        arrDone(164) = "False"
        arrValues(164) = "469F"
        arrDone(165) = "False"
        arrValues(165) = "048F"
        arrDone(166) = "False"
        arrValues(166) = "069"
        arrDone(167) = "True"
        arrValues(167) = "D"
        arrDone(168) = "True"
        arrValues(168) = "5"
        arrDone(169) = "False"
        arrValues(169) = "02689"
        arrDone(170) = "True"
        arrValues(170) = "1"
        arrDone(171) = "True"
        arrValues(171) = "B"
        arrDone(172) = "True"
        arrValues(172) = "E"
        arrDone(173) = "False"
        arrValues(173) = "02689"
        arrDone(174) = "False"
        arrValues(174) = "0689"
        arrDone(175) = "False"
        arrValues(175) = "024689"
        arrDone(176) = "False"
        arrValues(176) = "268B"
        arrDone(177) = "False"
        arrValues(177) = "068B"
        arrDone(178) = "False"
        arrValues(178) = "012689B"
        arrDone(179) = "True"
        arrValues(179) = "F"
        arrDone(180) = "False"
        arrValues(180) = "14569B"
        arrDone(181) = "True"
        arrValues(181) = "7"
        arrDone(182) = "True"
        arrValues(182) = "A"
        arrDone(183) = "True"
        arrValues(183) = "3"
        arrDone(184) = "False"
        arrValues(184) = "689E"
        arrDone(185) = "True"
        arrValues(185) = "D"
        arrDone(186) = "False"
        arrValues(186) = "24689E"
        arrDone(187) = "False"
        arrValues(187) = "0246E"
        arrDone(188) = "False"
        arrValues(188) = "24689B"
        arrDone(189) = "False"
        arrValues(189) = "012689"
        arrDone(190) = "True"
        arrValues(190) = "C"
        arrDone(191) = "False"
        arrValues(191) = "024689B"
        arrDone(192) = "False"
        arrValues(192) = "568AB"
        arrDone(193) = "True"
        arrValues(193) = "1"
        arrDone(194) = "False"
        arrValues(194) = "068AB"
        arrDone(195) = "False"
        arrValues(195) = "68AB"
        arrDone(196) = "True"
        arrValues(196) = "7"
        arrDone(197) = "True"
        arrValues(197) = "2"
        arrDone(198) = "True"
        arrValues(198) = "3"
        arrDone(199) = "False"
        arrValues(199) = "0589AF"
        arrDone(200) = "False"
        arrValues(200) = "689ACDEF"
        arrDone(201) = "False"
        arrValues(201) = "0689AEF"
        arrDone(202) = "False"
        arrValues(202) = "689ACDEF"
        arrDone(203) = "False"
        arrValues(203) = "06DE"
        arrDone(204) = "False"
        arrValues(204) = "689BCD"
        arrDone(205) = "False"
        arrValues(205) = "5689CD"
        arrDone(206) = "True"
        arrValues(206) = "4"
        arrDone(207) = "False"
        arrValues(207) = "689BCDF"
        arrDone(208) = "False"
        arrValues(208) = "5678A"
        arrDone(209) = "False"
        arrValues(209) = "568"
        arrDone(210) = "True"
        arrValues(210) = "3"
        arrDone(211) = "True"
        arrValues(211) = "D"
        arrDone(212) = "False"
        arrValues(212) = "159F"
        arrDone(213) = "False"
        arrValues(213) = "158AF"
        arrDone(214) = "True"
        arrValues(214) = "4"
        arrDone(215) = "False"
        arrValues(215) = "1589AF"
        arrDone(216) = "False"
        arrValues(216) = "1689ACF"
        arrDone(217) = "True"
        arrValues(217) = "B"
        arrDone(218) = "False"
        arrValues(218) = "689ACF"
        arrDone(219) = "False"
        arrValues(219) = "16"
        arrDone(220) = "True"
        arrValues(220) = "0"
        arrDone(221) = "False"
        arrValues(221) = "56789C"
        arrDone(222) = "True"
        arrValues(222) = "2"
        arrDone(223) = "True"
        arrValues(223) = "E"
        arrDone(224) = "False"
        arrValues(224) = "268B"
        arrDone(225) = "True"
        arrValues(225) = "F"
        arrDone(226) = "False"
        arrValues(226) = "0268B"
        arrDone(227) = "True"
        arrValues(227) = "C"
        arrDone(228) = "True"
        arrValues(228) = "E"
        arrDone(229) = "False"
        arrValues(229) = "08B"
        arrDone(230) = "False"
        arrValues(230) = "09B"
        arrDone(231) = "False"
        arrValues(231) = "089"
        arrDone(232) = "False"
        arrValues(232) = "689D"
        arrDone(233) = "True"
        arrValues(233) = "7"
        arrDone(234) = "False"
        arrValues(234) = "24689D"
        arrDone(235) = "True"
        arrValues(235) = "5"
        arrDone(236) = "False"
        arrValues(236) = "689BD"
        arrDone(237) = "False"
        arrValues(237) = "3689D"
        arrDone(238) = "True"
        arrValues(238) = "A"
        arrDone(239) = "True"
        arrValues(239) = "1"
        arrDone(240) = "True"
        arrValues(240) = "9"
        arrDone(241) = "False"
        arrValues(241) = "058B"
        arrDone(242) = "True"
        arrValues(242) = "E"
        arrDone(243) = "True"
        arrValues(243) = "4"
        arrDone(244) = "True"
        arrValues(244) = "D"
        arrDone(245) = "False"
        arrValues(245) = "0158ABF"
        arrDone(246) = "True"
        arrValues(246) = "C"
        arrDone(247) = "True"
        arrValues(247) = "6"
        arrDone(248) = "False"
        arrValues(248) = "18AF"
        arrDone(249) = "False"
        arrValues(249) = "028AF"
        arrDone(250) = "False"
        arrValues(250) = "28AF"
        arrDone(251) = "False"
        arrValues(251) = "0123"
        arrDone(252) = "False"
        arrValues(252) = "78B"
        arrDone(253) = "False"
        arrValues(253) = "3578"
        arrDone(254) = "False"
        arrValues(254) = "358"
        arrDone(255) = "False"
        arrValues(255) = "38BF"

        RefreshPuzzle()

    End Sub
    Private Sub LoadSolution4()

        arrDone(0) = "False"
        arrValues(0) = "016AD"
        arrDone(1) = "True"
        arrValues(1) = "4"
        arrDone(2) = "False"
        arrValues(2) = "1367AC"
        arrDone(3) = "False"
        arrValues(3) = "0367AC"
        arrDone(4) = "False"
        arrValues(4) = "8CD"
        arrDone(5) = "False"
        arrValues(5) = "0CD"
        arrDone(6) = "False"
        arrValues(6) = "08CD"
        arrDone(7) = "True"
        arrValues(7) = "9"
        arrDone(8) = "True"
        arrValues(8) = "E"
        arrDone(9) = "False"
        arrValues(9) = "17"
        arrDone(10) = "True"
        arrValues(10) = "B"
        arrDone(11) = "False"
        arrValues(11) = "178"
        arrDone(12) = "True"
        arrValues(12) = "2"
        arrDone(13) = "True"
        arrValues(13) = "5"
        arrDone(14) = "False"
        arrValues(14) = "067C"
        arrDone(15) = "True"
        arrValues(15) = "F"
        arrDone(16) = "False"
        arrValues(16) = "029BF"
        arrDone(17) = "False"
        arrValues(17) = "029CF"
        arrDone(18) = "False"
        arrValues(18) = "2C"
        arrDone(19) = "True"
        arrValues(19) = "8"
        arrDone(20) = "False"
        arrValues(20) = "2BCEF"
        arrDone(21) = "True"
        arrValues(21) = "5"
        arrDone(22) = "True"
        arrValues(22) = "7"
        arrDone(23) = "True"
        arrValues(23) = "1"
        arrDone(24) = "False"
        arrValues(24) = "029F"
        arrDone(25) = "True"
        arrValues(25) = "A"
        arrDone(26) = "True"
        arrValues(26) = "6"
        arrDone(27) = "False"
        arrValues(27) = "249"
        arrDone(28) = "False"
        arrValues(28) = "049B"
        arrDone(29) = "True"
        arrValues(29) = "3"
        arrDone(30) = "True"
        arrValues(30) = "D"
        arrDone(31) = "False"
        arrValues(31) = "049C"
        arrDone(32) = "False"
        arrValues(32) = "0259BF"
        arrDone(33) = "False"
        arrValues(33) = "0279CF"
        arrDone(34) = "False"
        arrValues(34) = "257C"
        arrDone(35) = "True"
        arrValues(35) = "E"
        arrDone(36) = "False"
        arrValues(36) = "2BCF"
        arrDone(37) = "False"
        arrValues(37) = "0BCF"
        arrDone(38) = "True"
        arrValues(38) = "A"
        arrDone(39) = "True"
        arrValues(39) = "6"
        arrDone(40) = "False"
        arrValues(40) = "0259F"
        arrDone(41) = "False"
        arrValues(41) = "2579F"
        arrDone(42) = "True"
        arrValues(42) = "D"
        arrDone(43) = "True"
        arrValues(43) = "3"
        arrDone(44) = "True"
        arrValues(44) = "8"
        arrDone(45) = "False"
        arrValues(45) = "0479BC"
        arrDone(46) = "True"
        arrValues(46) = "1"
        arrDone(47) = "False"
        arrValues(47) = "0479C"
        arrDone(48) = "False"
        arrValues(48) = "012569BDF"
        arrDone(49) = "False"
        arrValues(49) = "02679DF"
        arrDone(50) = "False"
        arrValues(50) = "12567"
        arrDone(51) = "False"
        arrValues(51) = "02679"
        arrDone(52) = "False"
        arrValues(52) = "28BDF"
        arrDone(53) = "True"
        arrValues(53) = "3"
        arrDone(54) = "False"
        arrValues(54) = "0248BD"
        arrDone(55) = "False"
        arrValues(55) = "48BD"
        arrDone(56) = "True"
        arrValues(56) = "C"
        arrDone(57) = "False"
        arrValues(57) = "12579F"
        arrDone(58) = "False"
        arrValues(58) = "01458F"
        arrDone(59) = "False"
        arrValues(59) = "1245789"
        arrDone(60) = "True"
        arrValues(60) = "E"
        arrDone(61) = "False"
        arrValues(61) = "0479B"
        arrDone(62) = "True"
        arrValues(62) = "A"
        arrDone(63) = "False"
        arrValues(63) = "04679"
        arrDone(64) = "False"
        arrValues(64) = "0245AF"
        arrDone(65) = "True"
        arrValues(65) = "1"
        arrDone(66) = "True"
        arrValues(66) = "E"
        arrDone(67) = "False"
        arrValues(67) = "02347A"
        arrDone(68) = "False"
        arrValues(68) = "8ABD"
        arrDone(69) = "False"
        arrValues(69) = "0ABD"
        arrDone(70) = "False"
        arrValues(70) = "038BD"
        arrDone(71) = "False"
        arrValues(71) = "78ABD"
        arrDone(72) = "False"
        arrValues(72) = "25ADF"
        arrDone(73) = "False"
        arrValues(73) = "25BDF"
        arrDone(74) = "True"
        arrValues(74) = "9"
        arrDone(75) = "False"
        arrValues(75) = "258BCD"
        arrDone(76) = "True"
        arrValues(76) = "6"
        arrDone(77) = "False"
        arrValues(77) = "0478BC"
        arrDone(78) = "False"
        arrValues(78) = "04578BC"
        arrDone(79) = "False"
        arrValues(79) = "024578CD"
        arrDone(80) = "False"
        arrValues(80) = "025AF"
        arrDone(81) = "False"
        arrValues(81) = "02AF"
        arrDone(82) = "True"
        arrValues(82) = "B"
        arrDone(83) = "False"
        arrValues(83) = "02A"
        arrDone(84) = "True"
        arrValues(84) = "4"
        arrDone(85) = "False"
        arrValues(85) = "0AD"
        arrDone(86) = "True"
        arrValues(86) = "9"
        arrDone(87) = "False"
        arrValues(87) = "8AD"
        arrDone(88) = "False"
        arrValues(88) = "25ADF"
        arrDone(89) = "False"
        arrValues(89) = "125DF"
        arrDone(90) = "True"
        arrValues(90) = "7"
        arrDone(91) = "True"
        arrValues(91) = "6"
        arrDone(92) = "True"
        arrValues(92) = "3"
        arrDone(93) = "True"
        arrValues(93) = "E"
        arrDone(94) = "False"
        arrValues(94) = "058C"
        arrDone(95) = "False"
        arrValues(95) = "01258CD"
        arrDone(96) = "True"
        arrValues(96) = "C"
        arrDone(97) = "False"
        arrValues(97) = "2679A"
        arrDone(98) = "True"
        arrValues(98) = "D"
        arrDone(99) = "False"
        arrValues(99) = "24679A"
        arrDone(100) = "False"
        arrValues(100) = "168AB"
        arrDone(101) = "False"
        arrValues(101) = "6AB"
        arrDone(102) = "True"
        arrValues(102) = "E"
        arrDone(103) = "True"
        arrValues(103) = "F"
        arrDone(104) = "False"
        arrValues(104) = "25A"
        arrDone(105) = "True"
        arrValues(105) = "3"
        arrDone(106) = "False"
        arrValues(106) = "158A"
        arrDone(107) = "True"
        arrValues(107) = "0"
        arrDone(108) = "False"
        arrValues(108) = "1479B"
        arrDone(109) = "False"
        arrValues(109) = "4789B"
        arrDone(110) = "False"
        arrValues(110) = "45789B"
        arrDone(111) = "False"
        arrValues(111) = "1245789"
        arrDone(112) = "False"
        arrValues(112) = "069"
        arrDone(113) = "False"
        arrValues(113) = "03679"
        arrDone(114) = "True"
        arrValues(114) = "8"
        arrDone(115) = "False"
        arrValues(115) = "03679"
        arrDone(116) = "True"
        arrValues(116) = "5"
        arrDone(117) = "True"
        arrValues(117) = "2"
        arrDone(118) = "False"
        arrValues(118) = "0136BD"
        arrDone(119) = "True"
        arrValues(119) = "C"
        arrDone(120) = "True"
        arrValues(120) = "4"
        arrDone(121) = "False"
        arrValues(121) = "1BDE"
        arrDone(122) = "False"
        arrValues(122) = "1"
        arrDone(123) = "False"
        arrValues(123) = "1BDE"
        arrDone(124) = "False"
        arrValues(124) = "0179BD"
        arrDone(125) = "False"
        arrValues(125) = "079B"
        arrDone(126) = "True"
        arrValues(126) = "F"
        arrDone(127) = "True"
        arrValues(127) = "A"
        arrDone(128) = "True"
        arrValues(128) = "7"
        arrDone(129) = "True"
        arrValues(129) = "5"
        arrDone(130) = "False"
        arrValues(130) = "146AC"
        arrDone(131) = "False"
        arrValues(131) = "046AC"
        arrDone(132) = "False"
        arrValues(132) = "69ACE"
        arrDone(133) = "False"
        arrValues(133) = "69AC"
        arrDone(134) = "False"
        arrValues(134) = "46C"
        arrDone(135) = "True"
        arrValues(135) = "3"
        arrDone(136) = "True"
        arrValues(136) = "B"
        arrDone(137) = "False"
        arrValues(137) = "169E"
        arrDone(138) = "True"
        arrValues(138) = "2"
        arrDone(139) = "True"
        arrValues(139) = "F"
        arrDone(140) = "False"
        arrValues(140) = "049A"
        arrDone(141) = "True"
        arrValues(141) = "D"
        arrDone(142) = "False"
        arrValues(142) = "0489C"
        arrDone(143) = "False"
        arrValues(143) = "0489C"
        arrDone(144) = "False"
        arrValues(144) = "012ADE"
        arrDone(145) = "False"
        arrValues(145) = "023ACD"
        arrDone(146) = "False"
        arrValues(146) = "123AC"
        arrDone(147) = "False"
        arrValues(147) = "023AC"
        arrDone(148) = "True"
        arrValues(148) = "7"
        arrDone(149) = "False"
        arrValues(149) = "9ACDF"
        arrDone(150) = "True"
        arrValues(150) = "5"
        arrDone(151) = "False"
        arrValues(151) = "ADE"
        arrDone(152) = "True"
        arrValues(152) = "8"
        arrDone(153) = "True"
        arrValues(153) = "4"
        arrDone(154) = "False"
        arrValues(154) = "013"
        arrDone(155) = "False"
        arrValues(155) = "19DE"
        arrDone(156) = "False"
        arrValues(156) = "09AF"
        arrDone(157) = "True"
        arrValues(157) = "6"
        arrDone(158) = "False"
        arrValues(158) = "039C"
        arrDone(159) = "True"
        arrValues(159) = "B"
        arrDone(160) = "False"
        arrValues(160) = "246DE"
        arrDone(161) = "False"
        arrValues(161) = "236D"
        arrDone(162) = "True"
        arrValues(162) = "9"
        arrDone(163) = "True"
        arrValues(163) = "B"
        arrDone(164) = "True"
        arrValues(164) = "0"
        arrDone(165) = "True"
        arrValues(165) = "8"
        arrDone(166) = "False"
        arrValues(166) = "246D"
        arrDone(167) = "False"
        arrValues(167) = "4DE"
        arrDone(168) = "False"
        arrValues(168) = "35D"
        arrDone(169) = "True"
        arrValues(169) = "C"
        arrDone(170) = "False"
        arrValues(170) = "35"
        arrDone(171) = "True"
        arrValues(171) = "A"
        arrDone(172) = "False"
        arrValues(172) = "47F"
        arrDone(173) = "True"
        arrValues(173) = "1"
        arrDone(174) = "False"
        arrValues(174) = "3457"
        arrDone(175) = "False"
        arrValues(175) = "3457"
        arrDone(176) = "False"
        arrValues(176) = "0468AD"
        arrDone(177) = "False"
        arrValues(177) = "036ACD"
        arrDone(178) = "False"
        arrValues(178) = "346AC"
        arrDone(179) = "True"
        arrValues(179) = "F"
        arrDone(180) = "False"
        arrValues(180) = "69ABCD"
        arrDone(181) = "True"
        arrValues(181) = "1"
        arrDone(182) = "False"
        arrValues(182) = "46BCD"
        arrDone(183) = "False"
        arrValues(183) = "4ABD"
        arrDone(184) = "False"
        arrValues(184) = "0359D"
        arrDone(185) = "False"
        arrValues(185) = "5679D"
        arrDone(186) = "False"
        arrValues(186) = "035"
        arrDone(187) = "False"
        arrValues(187) = "579D"
        arrDone(188) = "False"
        arrValues(188) = "0479A"
        arrDone(189) = "True"
        arrValues(189) = "2"
        arrDone(190) = "True"
        arrValues(190) = "E"
        arrDone(191) = "False"
        arrValues(191) = "0345789C"
        arrDone(192) = "False"
        arrValues(192) = "469A"
        arrDone(193) = "True"
        arrValues(193) = "8"
        arrDone(194) = "False"
        arrValues(194) = "467AC"
        arrDone(195) = "True"
        arrValues(195) = "5"
        arrDone(196) = "False"
        arrValues(196) = "169ABCD"
        arrDone(197) = "False"
        arrValues(197) = "69ABCD"
        arrDone(198) = "False"
        arrValues(198) = "16BCD"
        arrDone(199) = "True"
        arrValues(199) = "2"
        arrDone(200) = "False"
        arrValues(200) = "39ADF"
        arrDone(201) = "False"
        arrValues(201) = "9BDF"
        arrDone(202) = "True"
        arrValues(202) = "E"
        arrDone(203) = "False"
        arrValues(203) = "49BD"
        arrDone(204) = "False"
        arrValues(204) = "01479ABDF"
        arrDone(205) = "False"
        arrValues(205) = "0479ABF"
        arrDone(206) = "False"
        arrValues(206) = "034679B"
        arrDone(207) = "False"
        arrValues(207) = "0134679D"
        arrDone(208) = "False"
        arrValues(208) = "29A"
        arrDone(209) = "True"
        arrValues(209) = "E"
        arrDone(210) = "False"
        arrValues(210) = "27A"
        arrDone(211) = "True"
        arrValues(211) = "1"
        arrDone(212) = "True"
        arrValues(212) = "3"
        arrDone(213) = "True"
        arrValues(213) = "4"
        arrDone(214) = "False"
        arrValues(214) = "8BD"
        arrDone(215) = "False"
        arrValues(215) = "58ABD"
        arrDone(216) = "True"
        arrValues(216) = "6"
        arrDone(217) = "True"
        arrValues(217) = "0"
        arrDone(218) = "False"
        arrValues(218) = "5AF"
        arrDone(219) = "False"
        arrValues(219) = "259BD"
        arrDone(220) = "True"
        arrValues(220) = "C"
        arrDone(221) = "False"
        arrValues(221) = "789ABF"
        arrDone(222) = "False"
        arrValues(222) = "789B"
        arrDone(223) = "False"
        arrValues(223) = "789D"
        arrDone(224) = "False"
        arrValues(224) = "2469A"
        arrDone(225) = "True"
        arrValues(225) = "B"
        arrDone(226) = "True"
        arrValues(226) = "0"
        arrDone(227) = "False"
        arrValues(227) = "2469A"
        arrDone(228) = "False"
        arrValues(228) = "69AD"
        arrDone(229) = "True"
        arrValues(229) = "7"
        arrDone(230) = "True"
        arrValues(230) = "F"
        arrDone(231) = "False"
        arrValues(231) = "AD"
        arrDone(232) = "True"
        arrValues(232) = "1"
        arrDone(233) = "True"
        arrValues(233) = "8"
        arrDone(234) = "True"
        arrValues(234) = "C"
        arrDone(235) = "False"
        arrValues(235) = "249D"
        arrDone(236) = "True"
        arrValues(236) = "5"
        arrDone(237) = "False"
        arrValues(237) = "49A"
        arrDone(238) = "False"
        arrValues(238) = "3469"
        arrDone(239) = "False"
        arrValues(239) = "3469DE"
        arrDone(240) = "True"
        arrValues(240) = "3"
        arrDone(241) = "False"
        arrValues(241) = "69AC"
        arrDone(242) = "True"
        arrValues(242) = "F"
        arrDone(243) = "True"
        arrValues(243) = "D"
        arrDone(244) = "False"
        arrValues(244) = "1689ABC"
        arrDone(245) = "True"
        arrValues(245) = "E"
        arrDone(246) = "False"
        arrValues(246) = "168BC"
        arrDone(247) = "True"
        arrValues(247) = "0"
        arrDone(248) = "True"
        arrValues(248) = "7"
        arrDone(249) = "False"
        arrValues(249) = "59B"
        arrDone(250) = "False"
        arrValues(250) = "45A"
        arrDone(251) = "False"
        arrValues(251) = "459B"
        arrDone(252) = "False"
        arrValues(252) = "149AB"
        arrDone(253) = "False"
        arrValues(253) = "489AB"
        arrDone(254) = "True"
        arrValues(254) = "2"
        arrDone(255) = "False"
        arrValues(255) = "14689"

        RefreshPuzzle()

    End Sub
    Private Sub ButtonLoadSolution5()
        arrDone(0) = "True"
        arrValues(0) = "2"
        arrDone(1) = "True"
        arrValues(1) = "8"
        arrDone(2) = "True"
        arrValues(2) = "6"
        arrDone(3) = "True"
        arrValues(3) = "B"
        arrDone(4) = "True"
        arrValues(4) = "9"
        arrDone(5) = "True"
        arrValues(5) = "3"
        arrDone(6) = "True"
        arrValues(6) = "D"
        arrDone(7) = "True"
        arrValues(7) = "A"
        arrDone(8) = "True"
        arrValues(8) = "1"
        arrDone(9) = "True"
        arrValues(9) = "7"
        arrDone(10) = "True"
        arrValues(10) = "C"
        arrDone(11) = "True"
        arrValues(11) = "0"
        arrDone(12) = "True"
        arrValues(12) = "5"
        arrDone(13) = "True"
        arrValues(13) = "4"
        arrDone(14) = "True"
        arrValues(14) = "F"
        arrDone(15) = "True"
        arrValues(15) = "E"
        arrDone(16) = "True"
        arrValues(16) = "D"
        arrDone(17) = "True"
        arrValues(17) = "5"
        arrDone(18) = "True"
        arrValues(18) = "9"
        arrDone(19) = "True"
        arrValues(19) = "1"
        arrDone(20) = "True"
        arrValues(20) = "2"
        arrDone(21) = "True"
        arrValues(21) = "7"
        arrDone(22) = "True"
        arrValues(22) = "C"
        arrDone(23) = "True"
        arrValues(23) = "8"
        arrDone(24) = "True"
        arrValues(24) = "A"
        arrDone(25) = "True"
        arrValues(25) = "F"
        arrDone(26) = "True"
        arrValues(26) = "4"
        arrDone(27) = "True"
        arrValues(27) = "E"
        arrDone(28) = "True"
        arrValues(28) = "6"
        arrDone(29) = "True"
        arrValues(29) = "0"
        arrDone(30) = "True"
        arrValues(30) = "B"
        arrDone(31) = "True"
        arrValues(31) = "3"
        arrDone(32) = "True"
        arrValues(32) = "E"
        arrDone(33) = "True"
        arrValues(33) = "A"
        arrDone(34) = "True"
        arrValues(34) = "C"
        arrDone(35) = "True"
        arrValues(35) = "0"
        arrDone(36) = "True"
        arrValues(36) = "6"
        arrDone(37) = "True"
        arrValues(37) = "F"
        arrDone(38) = "True"
        arrValues(38) = "4"
        arrDone(39) = "True"
        arrValues(39) = "5"
        arrDone(40) = "True"
        arrValues(40) = "2"
        arrDone(41) = "True"
        arrValues(41) = "8"
        arrDone(42) = "True"
        arrValues(42) = "3"
        arrDone(43) = "True"
        arrValues(43) = "B"
        arrDone(44) = "True"
        arrValues(44) = "D"
        arrDone(45) = "True"
        arrValues(45) = "9"
        arrDone(46) = "True"
        arrValues(46) = "1"
        arrDone(47) = "True"
        arrValues(47) = "7"
        arrDone(48) = "True"
        arrValues(48) = "F"
        arrDone(49) = "True"
        arrValues(49) = "4"
        arrDone(50) = "True"
        arrValues(50) = "7"
        arrDone(51) = "True"
        arrValues(51) = "3"
        arrDone(52) = "True"
        arrValues(52) = "1"
        arrDone(53) = "True"
        arrValues(53) = "0"
        arrDone(54) = "True"
        arrValues(54) = "E"
        arrDone(55) = "True"
        arrValues(55) = "B"
        arrDone(56) = "True"
        arrValues(56) = "9"
        arrDone(57) = "True"
        arrValues(57) = "5"
        arrDone(58) = "True"
        arrValues(58) = "6"
        arrDone(59) = "True"
        arrValues(59) = "D"
        arrDone(60) = "True"
        arrValues(60) = "2"
        arrDone(61) = "True"
        arrValues(61) = "C"
        arrDone(62) = "True"
        arrValues(62) = "8"
        arrDone(63) = "True"
        arrValues(63) = "A"
        arrDone(64) = "True"
        arrValues(64) = "8"
        arrDone(65) = "True"
        arrValues(65) = "0"
        arrDone(66) = "True"
        arrValues(66) = "2"
        arrDone(67) = "True"
        arrValues(67) = "D"
        arrDone(68) = "True"
        arrValues(68) = "3"
        arrDone(69) = "True"
        arrValues(69) = "5"
        arrDone(70) = "True"
        arrValues(70) = "A"
        arrDone(71) = "True"
        arrValues(71) = "F"
        arrDone(72) = "True"
        arrValues(72) = "B"
        arrDone(73) = "True"
        arrValues(73) = "C"
        arrDone(74) = "True"
        arrValues(74) = "1"
        arrDone(75) = "True"
        arrValues(75) = "4"
        arrDone(76) = "True"
        arrValues(76) = "E"
        arrDone(77) = "True"
        arrValues(77) = "6"
        arrDone(78) = "True"
        arrValues(78) = "7"
        arrDone(79) = "True"
        arrValues(79) = "9"
        arrDone(80) = "True"
        arrValues(80) = "B"
        arrDone(81) = "True"
        arrValues(81) = "E"
        arrDone(82) = "True"
        arrValues(82) = "5"
        arrDone(83) = "True"
        arrValues(83) = "F"
        arrDone(84) = "True"
        arrValues(84) = "0"
        arrDone(85) = "True"
        arrValues(85) = "C"
        arrDone(86) = "True"
        arrValues(86) = "8"
        arrDone(87) = "True"
        arrValues(87) = "4"
        arrDone(88) = "True"
        arrValues(88) = "6"
        arrDone(89) = "True"
        arrValues(89) = "D"
        arrDone(90) = "True"
        arrValues(90) = "7"
        arrDone(91) = "True"
        arrValues(91) = "9"
        arrDone(92) = "True"
        arrValues(92) = "1"
        arrDone(93) = "True"
        arrValues(93) = "A"
        arrDone(94) = "True"
        arrValues(94) = "3"
        arrDone(95) = "True"
        arrValues(95) = "2"
        arrDone(96) = "True"
        arrValues(96) = "4"
        arrDone(97) = "True"
        arrValues(97) = "1"
        arrDone(98) = "True"
        arrValues(98) = "A"
        arrDone(99) = "True"
        arrValues(99) = "7"
        arrDone(100) = "True"
        arrValues(100) = "D"
        arrDone(101) = "True"
        arrValues(101) = "E"
        arrDone(102) = "True"
        arrValues(102) = "9"
        arrDone(103) = "True"
        arrValues(103) = "6"
        arrDone(104) = "True"
        arrValues(104) = "5"
        arrDone(105) = "True"
        arrValues(105) = "2"
        arrDone(106) = "True"
        arrValues(106) = "F"
        arrDone(107) = "True"
        arrValues(107) = "3"
        arrDone(108) = "True"
        arrValues(108) = "B"
        arrDone(109) = "True"
        arrValues(109) = "8"
        arrDone(110) = "True"
        arrValues(110) = "0"
        arrDone(111) = "True"
        arrValues(111) = "C"
        arrDone(112) = "True"
        arrValues(112) = "9"
        arrDone(113) = "True"
        arrValues(113) = "6"
        arrDone(114) = "True"
        arrValues(114) = "3"
        arrDone(115) = "True"
        arrValues(115) = "C"
        arrDone(116) = "True"
        arrValues(116) = "B"
        arrDone(117) = "True"
        arrValues(117) = "1"
        arrDone(118) = "True"
        arrValues(118) = "7"
        arrDone(119) = "True"
        arrValues(119) = "2"
        arrDone(120) = "True"
        arrValues(120) = "E"
        arrDone(121) = "True"
        arrValues(121) = "A"
        arrDone(122) = "True"
        arrValues(122) = "0"
        arrDone(123) = "True"
        arrValues(123) = "8"
        arrDone(124) = "True"
        arrValues(124) = "4"
        arrDone(125) = "True"
        arrValues(125) = "F"
        arrDone(126) = "True"
        arrValues(126) = "5"
        arrDone(127) = "True"
        arrValues(127) = "D"
        arrDone(128) = "True"
        arrValues(128) = "C"
        arrDone(129) = "True"
        arrValues(129) = "2"
        arrDone(130) = "True"
        arrValues(130) = "D"
        arrDone(131) = "True"
        arrValues(131) = "6"
        arrDone(132) = "True"
        arrValues(132) = "F"
        arrDone(133) = "True"
        arrValues(133) = "4"
        arrDone(134) = "True"
        arrValues(134) = "B"
        arrDone(135) = "True"
        arrValues(135) = "7"
        arrDone(136) = "True"
        arrValues(136) = "0"
        arrDone(137) = "True"
        arrValues(137) = "3"
        arrDone(138) = "True"
        arrValues(138) = "5"
        arrDone(139) = "True"
        arrValues(139) = "1"
        arrDone(140) = "True"
        arrValues(140) = "A"
        arrDone(141) = "True"
        arrValues(141) = "E"
        arrDone(142) = "True"
        arrValues(142) = "9"
        arrDone(143) = "True"
        arrValues(143) = "8"
        arrDone(144) = "True"
        arrValues(144) = "1"
        arrDone(145) = "True"
        arrValues(145) = "F"
        arrDone(146) = "True"
        arrValues(146) = "E"
        arrDone(147) = "True"
        arrValues(147) = "9"
        arrDone(148) = "True"
        arrValues(148) = "8"
        arrDone(149) = "True"
        arrValues(149) = "A"
        arrDone(150) = "True"
        arrValues(150) = "0"
        arrDone(151) = "True"
        arrValues(151) = "D"
        arrDone(152) = "True"
        arrValues(152) = "7"
        arrDone(153) = "True"
        arrValues(153) = "4"
        arrDone(154) = "True"
        arrValues(154) = "2"
        arrDone(155) = "True"
        arrValues(155) = "C"
        arrDone(156) = "True"
        arrValues(156) = "3"
        arrDone(157) = "True"
        arrValues(157) = "5"
        arrDone(158) = "True"
        arrValues(158) = "6"
        arrDone(159) = "True"
        arrValues(159) = "B"
        arrDone(160) = "True"
        arrValues(160) = "A"
        arrDone(161) = "True"
        arrValues(161) = "B"
        arrDone(162) = "True"
        arrValues(162) = "0"
        arrDone(163) = "True"
        arrValues(163) = "5"
        arrDone(164) = "True"
        arrValues(164) = "E"
        arrDone(165) = "True"
        arrValues(165) = "2"
        arrDone(166) = "True"
        arrValues(166) = "1"
        arrDone(167) = "True"
        arrValues(167) = "3"
        arrDone(168) = "True"
        arrValues(168) = "8"
        arrDone(169) = "True"
        arrValues(169) = "9"
        arrDone(170) = "True"
        arrValues(170) = "D"
        arrDone(171) = "True"
        arrValues(171) = "6"
        arrDone(172) = "True"
        arrValues(172) = "C"
        arrDone(173) = "True"
        arrValues(173) = "7"
        arrDone(174) = "True"
        arrValues(174) = "4"
        arrDone(175) = "True"
        arrValues(175) = "F"
        arrDone(176) = "True"
        arrValues(176) = "7"
        arrDone(177) = "True"
        arrValues(177) = "3"
        arrDone(178) = "True"
        arrValues(178) = "4"
        arrDone(179) = "True"
        arrValues(179) = "8"
        arrDone(180) = "True"
        arrValues(180) = "5"
        arrDone(181) = "True"
        arrValues(181) = "9"
        arrDone(182) = "True"
        arrValues(182) = "6"
        arrDone(183) = "True"
        arrValues(183) = "C"
        arrDone(184) = "True"
        arrValues(184) = "F"
        arrDone(185) = "True"
        arrValues(185) = "B"
        arrDone(186) = "True"
        arrValues(186) = "E"
        arrDone(187) = "True"
        arrValues(187) = "A"
        arrDone(188) = "True"
        arrValues(188) = "0"
        arrDone(189) = "True"
        arrValues(189) = "2"
        arrDone(190) = "True"
        arrValues(190) = "D"
        arrDone(191) = "True"
        arrValues(191) = "1"
        arrDone(192) = "True"
        arrValues(192) = "5"
        arrDone(193) = "True"
        arrValues(193) = "9"
        arrDone(194) = "True"
        arrValues(194) = "1"
        arrDone(195) = "True"
        arrValues(195) = "4"
        arrDone(196) = "True"
        arrValues(196) = "C"
        arrDone(197) = "True"
        arrValues(197) = "B"
        arrDone(198) = "True"
        arrValues(198) = "F"
        arrDone(199) = "True"
        arrValues(199) = "E"
        arrDone(200) = "True"
        arrValues(200) = "3"
        arrDone(201) = "True"
        arrValues(201) = "0"
        arrDone(202) = "True"
        arrValues(202) = "8"
        arrDone(203) = "True"
        arrValues(203) = "2"
        arrDone(204) = "True"
        arrValues(204) = "7"
        arrDone(205) = "True"
        arrValues(205) = "D"
        arrDone(206) = "True"
        arrValues(206) = "A"
        arrDone(207) = "True"
        arrValues(207) = "6"
        arrDone(208) = "True"
        arrValues(208) = "0"
        arrDone(209) = "True"
        arrValues(209) = "7"
        arrDone(210) = "True"
        arrValues(210) = "B"
        arrDone(211) = "True"
        arrValues(211) = "2"
        arrDone(212) = "True"
        arrValues(212) = "4"
        arrDone(213) = "True"
        arrValues(213) = "6"
        arrDone(214) = "True"
        arrValues(214) = "3"
        arrDone(215) = "True"
        arrValues(215) = "9"
        arrDone(216) = "True"
        arrValues(216) = "D"
        arrDone(217) = "True"
        arrValues(217) = "E"
        arrDone(218) = "True"
        arrValues(218) = "A"
        arrDone(219) = "True"
        arrValues(219) = "F"
        arrDone(220) = "True"
        arrValues(220) = "8"
        arrDone(221) = "True"
        arrValues(221) = "1"
        arrDone(222) = "True"
        arrValues(222) = "C"
        arrDone(223) = "True"
        arrValues(223) = "5"
        arrDone(224) = "True"
        arrValues(224) = "3"
        arrDone(225) = "True"
        arrValues(225) = "C"
        arrDone(226) = "True"
        arrValues(226) = "8"
        arrDone(227) = "True"
        arrValues(227) = "E"
        arrDone(228) = "True"
        arrValues(228) = "A"
        arrDone(229) = "True"
        arrValues(229) = "D"
        arrDone(230) = "True"
        arrValues(230) = "5"
        arrDone(231) = "True"
        arrValues(231) = "1"
        arrDone(232) = "True"
        arrValues(232) = "4"
        arrDone(233) = "True"
        arrValues(233) = "6"
        arrDone(234) = "True"
        arrValues(234) = "9"
        arrDone(235) = "True"
        arrValues(235) = "7"
        arrDone(236) = "True"
        arrValues(236) = "F"
        arrDone(237) = "True"
        arrValues(237) = "B"
        arrDone(238) = "True"
        arrValues(238) = "2"
        arrDone(239) = "True"
        arrValues(239) = "0"
        arrDone(240) = "True"
        arrValues(240) = "6"
        arrDone(241) = "True"
        arrValues(241) = "D"
        arrDone(242) = "True"
        arrValues(242) = "F"
        arrDone(243) = "True"
        arrValues(243) = "A"
        arrDone(244) = "True"
        arrValues(244) = "7"
        arrDone(245) = "True"
        arrValues(245) = "8"
        arrDone(246) = "True"
        arrValues(246) = "2"
        arrDone(247) = "True"
        arrValues(247) = "0"
        arrDone(248) = "True"
        arrValues(248) = "C"
        arrDone(249) = "True"
        arrValues(249) = "1"
        arrDone(250) = "True"
        arrValues(250) = "B"
        arrDone(251) = "True"
        arrValues(251) = "5"
        arrDone(252) = "True"
        arrValues(252) = "9"
        arrDone(253) = "True"
        arrValues(253) = "3"
        arrDone(254) = "True"
        arrValues(254) = "E"
        arrDone(255) = "True"
        arrValues(255) = "4"

        RefreshPuzzle()
    End Sub

    Private Sub ButtonClearTextbox1_Click(sender As Object, e As EventArgs) Handles ButtonClearTextbox1.Click
        Me.TextBox1.Text = ""
    End Sub

    Private Sub ButtonRemoveANumber_Click(sender As Object, e As EventArgs) Handles ButtonRemoveANumber.Click
        bRemoveANumberFromACell = True
    End Sub

    Private Sub ButtonUndoLast_Click(sender As Object, e As EventArgs) Handles ButtonUndoLast.Click
        Dim i As Int16

        For i = 0 To 255
            arrDone(i) = arrDone0(i)
            arrValues(i) = arrValues0(i)
        Next i

        RefreshPuzzle()
    End Sub

    Private Sub SaveCurrentSolution()
        Dim i As Int16

        For i = 0 To 255
            arrDone0(i) = arrDone(i)
            arrValues0(i) = arrValues(i)
        Next i

        RefreshPuzzle()
    End Sub

    'Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs)

    'End Sub

    'Private Sub ButtonOpenFile_Click(sender As Object, e As EventArgs) Handles ButtonOpenFile.Click
    '    'Me.TextBox1.Text = Me.OpenFileDialog1.ShowDialog()
    '    Dim i As Int16
    '    Dim n As Int16

    '    Dim strSolution As String
    '    Dim strTemp As String

    '    strSolution = Me.TextBox1.Text

    '    For i = 0 To 255
    '        If InStr(strSolution, "True") Then
    '            arrDone(i) = "True"
    '            n = InStr(strSolution, "True")
    '            strSolution = Mid(strSolution, n + 22)
    '        ElseIf InStr(strSolution, "False") Then
    '            arrDone(i) = "False"
    '            n = InStr(strSolution, "False")
    '            strSolution = Mid(strSolution, n + 22)
    '        End If
    '        If InStr(strSolution, Chr(34)) Then
    '            n = InStr(strSolution, Chr(34))
    '            strSolution = Mid(strSolution, n + 1)

    '        End If

    '        strSolution = Mid(strSolution, n + 22)
    '        strTemp = strSolution
    '    Next

    'End Sub

    Private Sub ButtonSolveAll_Click_1(sender As Object, e As EventArgs) Handles ButtonSolveAll.Click
        'Dim i As Int16
        Dim ROld As Int16

        For i = 0 To 1
            ROld = R

            If R > 0 Then
                CheckForOnlyNumberInARow()
                CheckForOnlyNumberInAColumn()
                CheckForOnlyNumberInASquare()
                CheckForOnlyOneNumber()
                'ROld = R
            End If
            'If ROld = R Then
            'Exit Sub
            'End If

        Next
    End Sub

    'Private Sub ButtonShowRemainders_Click(sender As Object, e As EventArgs) Handles ButtonShowRemainders.Click
    'Private Sub ButtonShowRemainers_Click(sender As Object, e As EventArgs)

    'End Sub

    Private Sub ButtonClearTextbox1_Click_1(sender As Object, e As EventArgs) Handles ButtonClearTextbox1.Click
        Me.TextBox1.Text = ""
    End Sub

    Private Sub btnLoadFunctions_Click(sender As Object, e As EventArgs) Handles btnLoadFunctions.Click
        'Me.TextBox1.Text = Me.TextBox1.Text & "Entering btnLoadFunctions_Click " & vbCrLf

        'Set this flag here to ensure it is set
        bShowRemainders = True
        bInitialShowRemainders = True

        StartingString = "0123456789ABCDEF"
        Dim i As Int16

        bShow = True
        bRemoveANumberFromACell = False

        'Initialize arrays arrDone to False and arrValues to the starting string
        For i = 0 To 255
            arrDone(i) = False
            arrValues(i) = StartingString
            'arrReferences(i, 3) = StartingString   'an alternate way to store all information (except done) in one array
        Next i

        arrReferences(0, 0) = 1
        Me.TextBox1.Text = Me.TextBox1.Text & "arrReferences(0, 0) is " & arrReferences(0, 0) & vbCrLf
        arrReferences(0, 1) = 1
        arrReferences(0, 2) = 1
        arrReferences(1, 0) = 1
        arrReferences(1, 1) = 2
        arrReferences(1, 2) = 1
        arrReferences(2, 0) = 1
        arrReferences(2, 1) = 3
        arrReferences(2, 2) = 1
        arrReferences(3, 0) = 1
        arrReferences(3, 1) = 4
        arrReferences(3, 2) = 1
        arrReferences(4, 0) = 1
        arrReferences(4, 1) = 5
        arrReferences(4, 2) = 2
        arrReferences(5, 0) = 1
        arrReferences(5, 1) = 6
        arrReferences(5, 2) = 2
        arrReferences(6, 0) = 1
        arrReferences(6, 1) = 7
        arrReferences(6, 2) = 2
        arrReferences(7, 0) = 1
        arrReferences(7, 1) = 7
        arrReferences(7, 2) = 2
        arrReferences(8, 0) = 1
        arrReferences(8, 1) = 9
        arrReferences(8, 2) = 3
        arrReferences(9, 0) = 1
        arrReferences(9, 1) = 10
        arrReferences(9, 2) = 3
        arrReferences(10, 0) = 1
        arrReferences(10, 1) = 11
        arrReferences(10, 2) = 3
        arrReferences(11, 0) = 1
        arrReferences(11, 1) = 12
        arrReferences(11, 2) = 3
        arrReferences(12, 0) = 1
        arrReferences(12, 1) = 13
        arrReferences(12, 2) = 4
        arrReferences(13, 0) = 1
        arrReferences(13, 1) = 14
        arrReferences(13, 2) = 4
        arrReferences(14, 0) = 1
        arrReferences(14, 1) = 15
        arrReferences(14, 2) = 4
        arrReferences(15, 0) = 1
        arrReferences(15, 1) = 16
        arrReferences(15, 2) = 4
        arrReferences(16, 0) = "2"
        arrReferences(16, 1) = "1"
        arrReferences(16, 2) = "1"
        arrReferences(17, 0) = "2"
        arrReferences(17, 1) = "2"
        arrReferences(17, 2) = "1"
        arrReferences(18, 0) = "2"
        arrReferences(18, 1) = "3"
        arrReferences(18, 2) = "1"
        arrReferences(19, 0) = "2"
        arrReferences(19, 1) = "4"
        arrReferences(19, 2) = "1"
        arrReferences(20, 0) = "2"
        arrReferences(20, 1) = "5"
        arrReferences(20, 2) = "2"
        arrReferences(21, 0) = "2"
        arrReferences(21, 1) = "6"
        arrReferences(21, 2) = "2"
        arrReferences(22, 0) = "2"
        arrReferences(22, 1) = "7"
        arrReferences(22, 2) = "2"
        arrReferences(23, 0) = "2"
        arrReferences(23, 1) = "8"
        arrReferences(23, 2) = "2"
        arrReferences(24, 0) = "2"
        arrReferences(24, 1) = "9"
        arrReferences(24, 2) = "3"
        arrReferences(25, 0) = "2"
        arrReferences(25, 1) = "10"
        arrReferences(25, 2) = "3"
        arrReferences(26, 0) = "2"
        arrReferences(26, 1) = "11"
        arrReferences(26, 2) = "3"
        arrReferences(27, 0) = "2"
        arrReferences(27, 1) = "12"
        arrReferences(27, 2) = "3"
        arrReferences(28, 0) = "2"
        arrReferences(28, 1) = "13"
        arrReferences(28, 2) = "4"
        arrReferences(29, 0) = "2"
        arrReferences(29, 1) = "14"
        arrReferences(29, 2) = "4"
        arrReferences(30, 0) = "2"
        arrReferences(30, 1) = "15"
        arrReferences(30, 2) = "4"
        arrReferences(31, 0) = "2"
        arrReferences(31, 1) = "16"
        arrReferences(31, 2) = "4"
        arrReferences(32, 0) = "3"
        arrReferences(32, 1) = "1"
        arrReferences(32, 2) = "1"
        arrReferences(33, 0) = "3"
        arrReferences(33, 1) = "2"
        arrReferences(33, 2) = "1"
        arrReferences(34, 0) = "3"
        arrReferences(34, 1) = "3"
        arrReferences(34, 2) = "1"
        arrReferences(35, 0) = "3"
        arrReferences(35, 1) = "4"
        arrReferences(35, 2) = "1"
        arrReferences(36, 0) = "3"
        arrReferences(36, 1) = "5"
        arrReferences(36, 2) = "2"
        arrReferences(37, 0) = "3"
        arrReferences(37, 1) = "6"
        arrReferences(37, 2) = "2"
        arrReferences(38, 0) = "3"
        arrReferences(38, 1) = "7"
        arrReferences(38, 2) = "2"
        arrReferences(39, 0) = "3"
        arrReferences(39, 1) = "8"
        arrReferences(39, 2) = "2"
        arrReferences(40, 0) = "3"
        arrReferences(40, 1) = "9"
        arrReferences(40, 2) = "3"
        arrReferences(41, 0) = "3"
        arrReferences(41, 1) = "10"
        arrReferences(41, 2) = "3"
        arrReferences(42, 0) = "3"
        arrReferences(42, 1) = "11"
        arrReferences(42, 2) = "3"
        arrReferences(43, 0) = "3"
        arrReferences(43, 1) = "12"
        arrReferences(43, 2) = "3"
        arrReferences(44, 0) = "3"
        arrReferences(44, 1) = "13"
        arrReferences(44, 2) = "4"
        arrReferences(45, 0) = "3"
        arrReferences(45, 1) = "14"
        arrReferences(45, 2) = "4"
        arrReferences(46, 0) = "3"
        arrReferences(46, 1) = "15"
        arrReferences(46, 2) = "4"
        arrReferences(47, 0) = "3"
        arrReferences(47, 1) = "16"
        arrReferences(47, 2) = "4"
        arrReferences(48, 0) = "4"
        arrReferences(48, 1) = "1"
        arrReferences(48, 2) = "1"
        arrReferences(49, 0) = "4"
        arrReferences(49, 1) = "2"
        arrReferences(49, 2) = "1"
        arrReferences(50, 0) = "4"
        arrReferences(50, 1) = "3"
        arrReferences(50, 2) = "1"
        arrReferences(51, 0) = "4"
        arrReferences(51, 1) = "4"
        arrReferences(51, 2) = "1"
        arrReferences(52, 0) = "4"
        arrReferences(52, 1) = "5"
        arrReferences(52, 2) = "2"
        arrReferences(53, 0) = "4"
        arrReferences(53, 1) = "6"
        arrReferences(53, 2) = "2"
        arrReferences(54, 0) = "4"
        arrReferences(54, 1) = "7"
        arrReferences(54, 2) = "2"
        arrReferences(55, 0) = "4"
        arrReferences(55, 1) = "8"
        arrReferences(55, 2) = "2"
        arrReferences(56, 0) = "4"
        arrReferences(56, 1) = "9"
        arrReferences(56, 2) = "3"
        arrReferences(57, 0) = "4"
        arrReferences(57, 1) = "10"
        arrReferences(57, 2) = "3"
        arrReferences(58, 0) = "4"
        arrReferences(58, 1) = "11"
        arrReferences(58, 2) = "3"
        arrReferences(59, 0) = "4"
        arrReferences(59, 1) = "12"
        arrReferences(59, 2) = "3"
        arrReferences(60, 0) = "4"
        arrReferences(60, 1) = "13"
        arrReferences(60, 2) = "4"
        arrReferences(61, 0) = "4"
        arrReferences(61, 1) = "14"
        arrReferences(61, 2) = "4"
        arrReferences(62, 0) = "4"
        arrReferences(62, 1) = "15"
        arrReferences(62, 2) = "4"
        arrReferences(63, 0) = "4"
        arrReferences(63, 1) = "16"
        arrReferences(63, 2) = "4"
        arrReferences(64, 0) = "5"
        arrReferences(64, 1) = "1"
        arrReferences(64, 2) = "5"
        arrReferences(65, 0) = "5"
        arrReferences(65, 1) = "2"
        arrReferences(65, 2) = "5"
        arrReferences(66, 0) = "5"
        arrReferences(66, 1) = "3"
        arrReferences(66, 2) = "5"
        arrReferences(67, 0) = "5"
        arrReferences(67, 1) = "4"
        arrReferences(67, 2) = "5"
        arrReferences(68, 0) = "5"
        arrReferences(68, 1) = "5"
        arrReferences(68, 2) = "6"
        arrReferences(69, 0) = "5"
        arrReferences(69, 1) = "6"
        arrReferences(69, 2) = "6"
        arrReferences(70, 0) = "5"
        arrReferences(70, 1) = "7"
        arrReferences(70, 2) = "6"
        arrReferences(71, 0) = "5"
        arrReferences(71, 1) = "8"
        arrReferences(71, 2) = "6"
        arrReferences(72, 0) = "5"
        arrReferences(72, 1) = "9"
        arrReferences(72, 2) = "7"
        arrReferences(73, 0) = "5"
        arrReferences(73, 1) = "10"
        arrReferences(73, 2) = "7"
        arrReferences(74, 0) = "5"
        arrReferences(74, 1) = "11"
        arrReferences(74, 2) = "7"
        arrReferences(75, 0) = "5"
        arrReferences(75, 1) = "12"
        arrReferences(75, 2) = "7"
        arrReferences(76, 0) = "5"
        arrReferences(76, 1) = "13"
        arrReferences(76, 2) = "8"
        arrReferences(77, 0) = "5"
        arrReferences(77, 1) = "14"
        arrReferences(77, 2) = "8"
        arrReferences(78, 0) = "5"
        arrReferences(78, 1) = "15"
        arrReferences(78, 2) = "8"
        arrReferences(79, 0) = "5"
        arrReferences(79, 1) = "16"
        arrReferences(79, 2) = "8"
        arrReferences(80, 0) = "6"
        arrReferences(80, 1) = "1"
        arrReferences(80, 2) = "5"
        arrReferences(81, 0) = "6"
        arrReferences(81, 1) = "2"
        arrReferences(81, 2) = "5"
        arrReferences(82, 0) = "6"
        arrReferences(82, 1) = "3"
        arrReferences(82, 2) = "5"
        arrReferences(83, 0) = "6"
        arrReferences(83, 1) = "4"
        arrReferences(83, 2) = "5"
        arrReferences(84, 0) = "6"
        arrReferences(84, 1) = "5"
        arrReferences(84, 2) = "6"
        arrReferences(85, 0) = "6"
        arrReferences(85, 1) = "6"
        arrReferences(85, 2) = "6"
        arrReferences(86, 0) = "6"
        arrReferences(86, 1) = "7"
        arrReferences(86, 2) = "6"
        arrReferences(87, 0) = "6"
        arrReferences(87, 1) = "8"
        arrReferences(87, 2) = "6"
        arrReferences(88, 0) = "6"
        arrReferences(88, 1) = "9"
        arrReferences(88, 2) = "7"
        arrReferences(89, 0) = "6"
        arrReferences(89, 1) = "10"
        arrReferences(89, 2) = "7"
        arrReferences(90, 0) = "6"
        arrReferences(90, 1) = "11"
        arrReferences(90, 2) = "7"
        arrReferences(91, 0) = "6"
        arrReferences(91, 1) = "12"
        arrReferences(91, 2) = "7"
        arrReferences(92, 0) = "6"
        arrReferences(92, 1) = "13"
        arrReferences(92, 2) = "8"
        arrReferences(93, 0) = "6"
        arrReferences(93, 1) = "14"
        arrReferences(93, 2) = "8"
        arrReferences(94, 0) = "6"
        arrReferences(94, 1) = "15"
        arrReferences(94, 2) = "8"
        arrReferences(95, 0) = "6"
        arrReferences(95, 1) = "16"
        arrReferences(95, 2) = "8"
        arrReferences(96, 0) = "7"
        arrReferences(96, 1) = "1"
        arrReferences(96, 2) = "5"
        arrReferences(97, 0) = "7"
        arrReferences(97, 1) = "2"
        arrReferences(97, 2) = "5"
        arrReferences(98, 0) = "7"
        arrReferences(98, 1) = "3"
        arrReferences(98, 2) = "5"
        arrReferences(99, 0) = "7"
        arrReferences(99, 1) = "4"
        arrReferences(99, 2) = "5"
        arrReferences(100, 0) = "7"
        arrReferences(100, 1) = "5"
        arrReferences(100, 2) = "6"
        arrReferences(101, 0) = "7"
        arrReferences(101, 1) = "6"
        arrReferences(101, 2) = "6"
        arrReferences(102, 0) = "7"
        arrReferences(102, 1) = "7"
        arrReferences(102, 2) = "6"
        arrReferences(103, 0) = "7"
        arrReferences(103, 1) = "8"
        arrReferences(103, 2) = "6"
        arrReferences(104, 0) = "7"
        arrReferences(104, 1) = "9"
        arrReferences(104, 2) = "7"
        arrReferences(105, 0) = "7"
        arrReferences(105, 1) = "10"
        arrReferences(105, 2) = "7"
        arrReferences(106, 0) = "7"
        arrReferences(106, 1) = "11"
        arrReferences(106, 2) = "7"
        arrReferences(107, 0) = "7"
        arrReferences(107, 1) = "12"
        arrReferences(107, 2) = "7"
        arrReferences(108, 0) = "7"
        arrReferences(108, 1) = "13"
        arrReferences(108, 2) = "8"
        arrReferences(109, 0) = "7"
        arrReferences(109, 1) = "14"
        arrReferences(109, 2) = "8"
        arrReferences(110, 0) = "7"
        arrReferences(110, 1) = "15"
        arrReferences(110, 2) = "8"
        arrReferences(111, 0) = "7"
        arrReferences(111, 1) = "16"
        arrReferences(111, 2) = "8"
        arrReferences(112, 0) = "8"
        arrReferences(112, 1) = "1"
        arrReferences(112, 2) = "5"
        arrReferences(113, 0) = "8"
        arrReferences(113, 1) = "2"
        arrReferences(113, 2) = "5"
        arrReferences(114, 0) = "8"
        arrReferences(114, 1) = "3"
        arrReferences(114, 2) = "5"
        arrReferences(115, 0) = "8"
        arrReferences(115, 1) = "4"
        arrReferences(115, 2) = "5"
        arrReferences(116, 0) = "8"
        arrReferences(116, 1) = "5"
        arrReferences(116, 2) = "6"
        arrReferences(117, 0) = "8"
        arrReferences(117, 1) = "6"
        arrReferences(117, 2) = "6"
        arrReferences(118, 0) = "8"
        arrReferences(118, 1) = "7"
        arrReferences(118, 2) = "6"
        arrReferences(119, 0) = "8"
        arrReferences(119, 1) = "8"
        arrReferences(119, 2) = "6"
        arrReferences(120, 0) = "8"
        arrReferences(120, 1) = "9"
        arrReferences(120, 2) = "7"
        arrReferences(121, 0) = "8"
        arrReferences(121, 1) = "10"
        arrReferences(121, 2) = "7"
        arrReferences(122, 0) = "8"
        arrReferences(122, 1) = "11"
        arrReferences(122, 2) = "7"
        arrReferences(123, 0) = "8"
        arrReferences(123, 1) = "12"
        arrReferences(123, 2) = "7"
        arrReferences(124, 0) = "8"
        arrReferences(124, 1) = "13"
        arrReferences(124, 2) = "8"
        arrReferences(125, 0) = "8"
        arrReferences(125, 1) = "14"
        arrReferences(125, 2) = "8"
        arrReferences(126, 0) = "8"
        arrReferences(126, 1) = "15"
        arrReferences(126, 2) = "8"
        arrReferences(127, 0) = "8"
        arrReferences(127, 1) = "16"
        arrReferences(127, 2) = "8"
        arrReferences(128, 0) = "9"
        arrReferences(128, 1) = "1"
        arrReferences(128, 2) = "9"
        arrReferences(129, 0) = "9"
        arrReferences(129, 1) = "2"
        arrReferences(129, 2) = "9"
        arrReferences(130, 0) = "9"
        arrReferences(130, 1) = "3"
        arrReferences(130, 2) = "9"
        arrReferences(131, 0) = "9"
        arrReferences(131, 1) = "4"
        arrReferences(131, 2) = "9"
        arrReferences(132, 0) = "9"
        arrReferences(132, 1) = "5"
        arrReferences(132, 2) = "10"
        arrReferences(133, 0) = "9"
        arrReferences(133, 1) = "6"
        arrReferences(133, 2) = "10"
        arrReferences(134, 0) = "9"
        arrReferences(134, 1) = "7"
        arrReferences(134, 2) = "10"
        arrReferences(135, 0) = "9"
        arrReferences(135, 1) = "8"
        arrReferences(135, 2) = "10"
        arrReferences(136, 0) = "9"
        arrReferences(136, 1) = "9"
        arrReferences(136, 2) = "11"
        arrReferences(137, 0) = "9"
        arrReferences(137, 1) = "10"
        arrReferences(137, 2) = "11"
        arrReferences(138, 0) = "9"
        arrReferences(138, 1) = "11"
        arrReferences(138, 2) = "11"
        arrReferences(139, 0) = "9"
        arrReferences(139, 1) = "12"
        arrReferences(139, 2) = "11"
        arrReferences(140, 0) = "9"
        arrReferences(140, 1) = "13"
        arrReferences(140, 2) = "12"
        arrReferences(141, 0) = "9"
        arrReferences(141, 1) = "14"
        arrReferences(141, 2) = "12"
        arrReferences(142, 0) = "9"
        arrReferences(142, 1) = "15"
        arrReferences(142, 2) = "12"
        arrReferences(143, 0) = "9"
        arrReferences(143, 1) = "16"
        arrReferences(143, 2) = "12"
        arrReferences(144, 0) = "10"
        arrReferences(144, 1) = "1"
        arrReferences(144, 2) = "9"
        arrReferences(145, 0) = "10"
        arrReferences(145, 1) = "2"
        arrReferences(145, 2) = "9"
        arrReferences(146, 0) = "10"
        arrReferences(146, 1) = "3"
        arrReferences(146, 2) = "9"
        arrReferences(147, 0) = "10"
        arrReferences(147, 1) = "4"
        arrReferences(147, 2) = "9"
        arrReferences(148, 0) = "10"
        arrReferences(148, 1) = "5"
        arrReferences(148, 2) = "10"
        arrReferences(149, 0) = "10"
        arrReferences(149, 1) = "6"
        arrReferences(149, 2) = "10"
        arrReferences(150, 0) = "10"
        arrReferences(150, 1) = "7"
        arrReferences(150, 2) = "10"
        arrReferences(151, 0) = "10"
        arrReferences(151, 1) = "8"
        arrReferences(151, 2) = "10"
        arrReferences(152, 0) = "10"
        arrReferences(152, 1) = "9"
        arrReferences(152, 2) = "11"
        arrReferences(153, 0) = "10"
        arrReferences(153, 1) = "10"
        arrReferences(153, 2) = "11"
        arrReferences(154, 0) = "10"
        arrReferences(154, 1) = "11"
        arrReferences(154, 2) = "11"
        arrReferences(155, 0) = "10"
        arrReferences(155, 1) = "12"
        arrReferences(155, 2) = "11"
        arrReferences(156, 0) = "10"
        arrReferences(156, 1) = "13"
        arrReferences(156, 2) = "12"
        arrReferences(157, 0) = "10"
        arrReferences(157, 1) = "14"
        arrReferences(157, 2) = "12"
        arrReferences(158, 0) = "10"
        arrReferences(158, 1) = "15"
        arrReferences(158, 2) = "12"
        arrReferences(159, 0) = "10"
        arrReferences(159, 1) = "16"
        arrReferences(159, 2) = "12"
        arrReferences(160, 0) = "11"
        arrReferences(160, 1) = "1"
        arrReferences(160, 2) = "9"
        arrReferences(161, 0) = "11"
        arrReferences(161, 1) = "2"
        arrReferences(161, 2) = "9"
        arrReferences(162, 0) = "11"
        arrReferences(162, 1) = "3"
        arrReferences(162, 2) = "9"
        arrReferences(163, 0) = "11"
        arrReferences(163, 1) = "4"
        arrReferences(163, 2) = "9"
        arrReferences(164, 0) = "11"
        arrReferences(164, 1) = "5"
        arrReferences(164, 2) = "10"
        arrReferences(165, 0) = "11"
        arrReferences(165, 1) = "6"
        arrReferences(165, 2) = "10"
        arrReferences(166, 0) = "11"
        arrReferences(166, 1) = "7"
        arrReferences(166, 2) = "10"
        arrReferences(167, 0) = "11"
        arrReferences(167, 1) = "8"
        arrReferences(167, 2) = "10"
        arrReferences(168, 0) = "11"
        arrReferences(168, 1) = "9"
        arrReferences(168, 2) = "11"
        arrReferences(169, 0) = "11"
        arrReferences(169, 1) = "10"
        arrReferences(169, 2) = "11"
        arrReferences(170, 0) = "11"
        arrReferences(170, 1) = "11"
        arrReferences(170, 2) = "11"
        arrReferences(171, 0) = "11"
        arrReferences(171, 1) = "12"
        arrReferences(171, 2) = "11"
        arrReferences(172, 0) = "11"
        arrReferences(172, 1) = "13"
        arrReferences(172, 2) = "12"
        arrReferences(173, 0) = "11"
        arrReferences(173, 1) = "14"
        arrReferences(173, 2) = "12"
        arrReferences(174, 0) = "11"
        arrReferences(174, 1) = "15"
        arrReferences(174, 2) = "12"
        arrReferences(175, 0) = "11"
        arrReferences(175, 1) = "16"
        arrReferences(175, 2) = "12"
        arrReferences(176, 0) = "12"
        arrReferences(176, 1) = "1"
        arrReferences(176, 2) = "9"
        arrReferences(177, 0) = "12"
        arrReferences(177, 1) = "2"
        arrReferences(177, 2) = "9"
        arrReferences(178, 0) = "12"
        arrReferences(178, 1) = "3"
        arrReferences(178, 2) = "9"
        arrReferences(179, 0) = "12"
        arrReferences(179, 1) = "4"
        arrReferences(179, 2) = "9"
        arrReferences(180, 0) = "12"
        arrReferences(180, 1) = "5"
        arrReferences(180, 2) = "10"
        arrReferences(181, 0) = "12"
        arrReferences(181, 1) = "6"
        arrReferences(181, 2) = "10"
        arrReferences(182, 0) = "12"
        arrReferences(182, 1) = "7"
        arrReferences(182, 2) = "10"
        arrReferences(183, 0) = "12"
        arrReferences(183, 1) = "8"
        arrReferences(183, 2) = "10"
        arrReferences(184, 0) = "12"
        arrReferences(184, 1) = "9"
        arrReferences(184, 2) = "11"
        arrReferences(185, 0) = "12"
        arrReferences(185, 1) = "10"
        arrReferences(185, 2) = "11"
        arrReferences(186, 0) = "12"
        arrReferences(186, 1) = "11"
        arrReferences(186, 2) = "11"
        arrReferences(187, 0) = "12"
        arrReferences(187, 1) = "12"
        arrReferences(187, 2) = "11"
        arrReferences(188, 0) = "12"
        arrReferences(188, 1) = "13"
        arrReferences(188, 2) = "12"
        arrReferences(189, 0) = "12"
        arrReferences(189, 1) = "14"
        arrReferences(189, 2) = "12"
        arrReferences(190, 0) = "12"
        arrReferences(190, 1) = "15"
        arrReferences(190, 2) = "12"
        arrReferences(191, 0) = "12"
        arrReferences(191, 1) = "16"
        arrReferences(191, 2) = "12"
        arrReferences(192, 0) = "13"
        arrReferences(192, 1) = "1"
        arrReferences(192, 2) = "13"
        arrReferences(193, 0) = "13"
        arrReferences(193, 1) = "2"
        arrReferences(193, 2) = "13"
        arrReferences(194, 0) = "13"
        arrReferences(194, 1) = "3"
        arrReferences(194, 2) = "13"
        arrReferences(195, 0) = "13"
        arrReferences(195, 1) = "4"
        arrReferences(195, 2) = "13"
        arrReferences(196, 0) = "13"
        arrReferences(196, 1) = "5"
        arrReferences(196, 2) = "14"
        arrReferences(197, 0) = "13"
        arrReferences(197, 1) = "6"
        arrReferences(197, 2) = "14"
        arrReferences(198, 0) = "13"
        arrReferences(198, 1) = "7"
        arrReferences(198, 2) = "14"
        arrReferences(199, 0) = "13"
        arrReferences(199, 1) = "8"
        arrReferences(199, 2) = "14"
        arrReferences(200, 0) = "13"
        arrReferences(200, 1) = "9"
        arrReferences(200, 2) = "15"
        arrReferences(201, 0) = "13"
        arrReferences(201, 1) = "10"
        arrReferences(201, 2) = "15"
        arrReferences(202, 0) = "13"
        arrReferences(202, 1) = "11"
        arrReferences(202, 2) = "15"
        arrReferences(203, 0) = "13"
        arrReferences(203, 1) = "12"
        arrReferences(203, 2) = "15"
        arrReferences(204, 0) = "13"
        arrReferences(204, 1) = "13"
        arrReferences(204, 2) = "16"
        arrReferences(205, 0) = "13"
        arrReferences(205, 1) = "14"
        arrReferences(205, 2) = "16"
        arrReferences(206, 0) = "13"
        arrReferences(206, 1) = "15"
        arrReferences(206, 2) = "16"
        arrReferences(207, 0) = "13"
        arrReferences(207, 1) = "16"
        arrReferences(207, 2) = "16"
        arrReferences(208, 0) = "14"
        arrReferences(208, 1) = "1"
        arrReferences(208, 2) = "13"
        arrReferences(209, 0) = "14"
        arrReferences(209, 1) = "2"
        arrReferences(209, 2) = "13"
        arrReferences(210, 0) = "14"
        arrReferences(210, 1) = "3"
        arrReferences(210, 2) = "13"
        arrReferences(211, 0) = "14"
        arrReferences(211, 1) = "4"
        arrReferences(211, 2) = "13"
        arrReferences(212, 0) = "14"
        arrReferences(212, 1) = "5"
        arrReferences(212, 2) = "14"
        arrReferences(213, 0) = "14"
        arrReferences(213, 1) = "6"
        arrReferences(213, 2) = "14"
        arrReferences(214, 0) = "14"
        arrReferences(214, 1) = "7"
        arrReferences(214, 2) = "14"
        arrReferences(215, 0) = "14"
        arrReferences(215, 1) = "8"
        arrReferences(215, 2) = "14"
        arrReferences(216, 0) = "14"
        arrReferences(216, 1) = "9"
        arrReferences(216, 2) = "15"
        arrReferences(217, 0) = "14"
        arrReferences(217, 1) = "10"
        arrReferences(217, 2) = "15"
        arrReferences(218, 0) = "14"
        arrReferences(218, 1) = "11"
        arrReferences(218, 2) = "15"
        arrReferences(219, 0) = "14"
        arrReferences(219, 1) = "12"
        arrReferences(219, 2) = "15"
        arrReferences(220, 0) = "14"
        arrReferences(220, 1) = "13"
        arrReferences(220, 2) = "16"
        arrReferences(221, 0) = "14"
        arrReferences(221, 1) = "14"
        arrReferences(221, 2) = "16"
        arrReferences(222, 0) = "14"
        arrReferences(222, 1) = "15"
        arrReferences(222, 2) = "16"
        arrReferences(223, 0) = "14"
        arrReferences(223, 1) = "16"
        arrReferences(223, 2) = "16"
        arrReferences(224, 0) = "15"
        arrReferences(224, 1) = "1"
        arrReferences(224, 2) = "13"
        arrReferences(225, 0) = "15"
        arrReferences(225, 1) = "2"
        arrReferences(225, 2) = "13"
        arrReferences(226, 0) = "15"
        arrReferences(226, 1) = "3"
        arrReferences(226, 2) = "13"
        arrReferences(227, 0) = "15"
        arrReferences(227, 1) = "4"
        arrReferences(227, 2) = "13"
        arrReferences(228, 0) = "15"
        arrReferences(228, 1) = "5"
        arrReferences(228, 2) = "14"
        arrReferences(229, 0) = "15"
        arrReferences(229, 1) = "6"
        arrReferences(229, 2) = "14"
        arrReferences(230, 0) = "15"
        arrReferences(230, 1) = "7"
        arrReferences(230, 2) = "14"
        arrReferences(231, 0) = "15"
        arrReferences(231, 1) = "8"
        arrReferences(231, 2) = "14"
        arrReferences(232, 0) = "15"
        arrReferences(232, 1) = "9"
        arrReferences(232, 2) = "15"
        arrReferences(233, 0) = "15"
        arrReferences(233, 1) = "10"
        arrReferences(233, 2) = "15"
        arrReferences(234, 0) = "15"
        arrReferences(234, 1) = "11"
        arrReferences(234, 2) = "15"
        arrReferences(235, 0) = "15"
        arrReferences(235, 1) = "12"
        arrReferences(235, 2) = "15"
        arrReferences(236, 0) = "15"
        arrReferences(236, 1) = "13"
        arrReferences(236, 2) = "16"
        arrReferences(237, 0) = "15"
        arrReferences(237, 1) = "14"
        arrReferences(237, 2) = "16"
        arrReferences(238, 0) = "15"
        arrReferences(238, 1) = "15"
        arrReferences(238, 2) = "16"
        arrReferences(239, 0) = "15"
        arrReferences(239, 1) = "16"
        arrReferences(239, 2) = "16"
        arrReferences(240, 0) = "16"
        arrReferences(240, 1) = "1"
        arrReferences(240, 2) = "13"
        arrReferences(241, 0) = "16"
        arrReferences(241, 1) = "2"
        arrReferences(241, 2) = "13"
        arrReferences(242, 0) = "16"
        arrReferences(242, 1) = "3"
        arrReferences(242, 2) = "13"
        arrReferences(243, 0) = "16"
        arrReferences(243, 1) = "4"
        arrReferences(243, 2) = "13"
        arrReferences(244, 0) = "16"
        arrReferences(244, 1) = "5"
        arrReferences(244, 2) = "14"
        arrReferences(245, 0) = "16"
        arrReferences(245, 1) = "6"
        arrReferences(245, 2) = "14"
        arrReferences(246, 0) = "16"
        arrReferences(246, 1) = "7"
        arrReferences(246, 2) = "14"
        arrReferences(247, 0) = "16"
        arrReferences(247, 1) = "8"
        arrReferences(247, 2) = "14"
        arrReferences(248, 0) = "16"
        arrReferences(248, 1) = "9"
        arrReferences(248, 2) = "15"
        arrReferences(249, 0) = "16"
        arrReferences(249, 1) = "10"
        arrReferences(249, 2) = "15"
        arrReferences(250, 0) = "16"
        arrReferences(250, 1) = "11"
        arrReferences(250, 2) = "15"
        arrReferences(251, 0) = "16"
        arrReferences(251, 1) = "12"
        arrReferences(251, 2) = "15"
        arrReferences(252, 0) = "16"
        arrReferences(252, 1) = "13"
        arrReferences(252, 2) = "16"
        arrReferences(253, 0) = "16"
        arrReferences(253, 1) = "14"
        arrReferences(253, 2) = "16"
        arrReferences(254, 0) = "16"
        arrReferences(254, 1) = "15"
        arrReferences(254, 2) = "16"
        arrReferences(255, 0) = "16"
        arrReferences(255, 1) = "16"
        arrReferences(255, 2) = "16"

        RefreshPuzzle()


    End Sub

    Private Sub ButtonShowRemainers_Click(sender As Object, e As EventArgs) Handles ButtonShowRemainers.Click
        Dim i As Int16

        'First, save the current solution
        If bInitialShowRemainders = True Then
            'SaveCurrentSolution()
            For i = 0 To 255
                'arrDone0(i) = arrDone(i)  'This wont change so don't save it
                arrValues0(i) = arrValues(i)
            Next i
            bInitialShowRemainders = False

        ElseIf bInitialShowRemainders = False Then
            If bShowRemainders = False Then
                For i = 0 To 255
                    If (arrDone(i)) = False Then
                        arrValues0(i) = arrValues(i)
                        arrValues(i) = NullValue
                    End If
                Next i
                bShowRemainders = True
            ElseIf bShowRemainders = True Then
                For i = 0 To 255
                    If (arrDone(i)) = False And arrValues(i) = NullValue Then
                        arrValues(i) = arrValues0(i)
                    End If
                Next i
                bShowRemainders = False

            End If
        End If

        RefreshPuzzle()

        'arrNullValues = ""

        'If bShowRemainders = True Then
        '    bShowRemainders = False
        'ElseIf bShowRemainders = False Then
        '    bShowRemainders = True
        'End If

        'SaveCurrentSolution()


        'RefreshPuzzle()

        'Now, Hide() Or Show() the values
        'If bShowRemainders = False Then
        '    For i = 0 To 255
        '        If (arrDone(i)) = False Then
        '            arrValues1(i) = arrValues(i)
        '            arrValues(i) = NullValue
        '        End If
        '    Next i
        '    bShowRemainders = True
        'ElseIf bShowRemainders = True Then
        '    'RestoreSolution1()
        '    For i = 0 To 255
        '        If (arrDone(i)) = False Then
        '            arrValues(i) = arrValues1(i)
        '        End If
        '    Next i
        '    bShowRemainders = False
        'End If

        'If bShowRemainders = True Then
        '    bShowRemainders = False
        'ElseIf bShowRemainders = False Then
        '    bShowRemainders = True
        'End If

        'For i = 0 To 255
        '    arrDone(i) = arrDone1(i)
        '    arrValues(i) = arrValues1(i)
        'Next i
        'RefreshPuzzle()

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox3.Text = Hints
        ElseIf CheckBox1.Checked = False Then
            TextBox3.Text = ""
        End If
    End Sub

    Private Sub frmSudokuMonster_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class