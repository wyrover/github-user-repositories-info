
Class clsThief


    Private value_ ' ��ȡ��������
    Private src_ ' Ҫ͵��Ŀ��URL��ַ
    Private isGet_ ' �ж��Ƿ��Ѿ�͵��
	Private cookie_ 

    ' ��ֵ��Ҫ͵��Ŀ��URL��ַ/����

    Public Property Let src(Str)
        src_ = Str
    End Property

    '����ֵ��������ȡ��Ӧ���෽���ӹ���������/����

    Public Property Get Value
        Value = value_
    End Property

	Public Property Get Cookie
        Cookie = cookie_
    End Property

    Public Property Get Version
        Version = "�ȷ溣���� Version 2004"
    End Property

    Private Sub class_initialize()
        value_ = ""
        src_ = ""
        isGet_ = False
    End Sub

    Private Sub class_terminate()
    End Sub

    ' ���Ĵ���

    Private Function BytesToBstr(body, Cset)
        Dim objstream
        Set objstream = CreateObject("adodb.stream")
        objstream.Type = 1
        objstream.Mode = 3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText
        objstream.Close
        Set objstream = Nothing
    End Function

    ' ��ȡĿ��URL��ַ��HTML����/����

    Public Sub steal(encode)
        If src_<>"" Then
            Dim Http
            Set Http = CreateObject("MSXML2.XMLHTTP")
            Http.Open "GET", src_ , false
			'Http.Open "GET", src_ , false
            Http.send()
			'cookie = Http.getResponseHeader("Set-Cookie")
            If Http.readystate<>4 Then
                Exit Sub
            End If
            value_ = BytesToBSTR(Http.responseBody, encode)
            isGet_ = True
            Set http = Nothing
            If Err.Number<>0 Then Err.Clear
        Else
            response.Write("<script>alert(""��������src���ԣ�"")</script>")
        End If
    End Sub

    ' ɾ��͵��������������Ļ��С��س����Ա��һ���ӹ�/����

    Public Sub noReturn()
        If isGet_ = false Then Call steal()
        value_ = Replace(Replace(value_ , vbCr, ""), vbLf, "")
    End Sub

    ' ��͵���������еĸ����ַ�������ֵ����/����
	' �����ֱ��Ǿ��ַ���,���ַ���
    Public Sub change(oldStr, Str) 
        If isGet_ = false Then Call steal()
        value_ = Replace(value_ , oldStr, Str)
    End Sub

    ' ��ָ����β�ַ�����͵ȡ�����ݽ��вü�����������β�ַ�����/����
    ' �����ֱ������ַ���,β�ַ���

    Public Sub cut(head, bot)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head) + Len(head), InStr(value_ , bot) - InStr(value_ , head) - Len(head))
    End Sub

    ' ��ָ����β�ַ�����͵ȡ�����ݽ��вü���������β�ַ�����/����
    ' �����ֱ������ַ���,β�ַ���

    Public Sub cutX(head, bot)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head), InStr(value_ , bot) - InStr(value_ , head) + Len(bot))
    End Sub

    '��ָ����β�ַ���λ��ƫ��ָ���͵ȡ�����ݽ��вü�/����
    '�����ֱ������ַ���,��ƫ��ֵ,β�ַ���,βƫ��ֵ,��ƫ���ø�ֵ,ƫ��ָ�뵥λΪ�ַ���

    Public Sub cutBy(head, headCusor, bot, botCusor)
        If isGet_ = false Then Call steal()
        value_ = Mid(value_ , InStr(value_ , head) + Len(head) + headCusor, InStr(value_ , bot) -1 + botCusor - InStr(value_ , head) - Len(head) - headcusor)
    End Sub

    '��ָ����β�ַ�����͵ȡ����������ֵ�����滻����������β�ַ�����/����
    '�����ֱ������ַ���,β�ַ���,��ֵ,��ֵλ����Ϊ����

    Public Sub filt(head, bot, Str)
        If isGet_ = false Then Call steal()
        value_ = Replace(value_, Mid(value_ , InStr(value_ , head) + Len(head), InStr(value_ , bot) -1), Str)
    End Sub

    '��ָ����β�ַ�����͵ȡ����������ֵ�����滻��������β�ַ�����/����
    '�����ֱ������ַ���,β�ַ���,��ֵ,��ֵΪ����Ϊ����

    Public Sub filtX(head, bot, Str)
        If isGet_ = false Then Call steal()
        value_ = Replace(value_, Mid(value_ , InStr(value_ , head), InStr(value_ , bot) + Len(bot) -1), Str)
    End Sub

    '��ָ����β�ַ���λ��ƫ��ָ���͵ȡ��������ֵ�����滻/����
    '�����ֱ������ַ���,��ƫ��ֵ,β�ַ���,βƫ��ֵ,��ֵ,��ƫ���ø�ֵ,ƫ��ָ�뵥λΪ�ַ���,��ֵΪ����Ϊ����

    Public Sub filtBy(head, headCusor, bot, botCusor, Str)

        If isGet_ = false Then Call steal()
        value_ = Replace(value_ , Mid(value_ , InStr(value_ , head) + Len(head) + headCusor, InStr(value_ , bot) -1 + botCusor - InStr(value_ , head) - Len(head) - headcusor), Str)
    End Sub

    '��͵ȡ�������еľ���URL��ַ��Ϊ������Ե�ַ

    Public Sub local()
        Dim tempReg
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg.Pattern = "^(http|https|ftp):(\/\/|////)(\w+.)+(com|net|org|cc|tv|cn|biz|com.cn|net.cn|sh.cn)\/"
        value_ = tempReg.Replace(value_ , "")
        Set tempReg = Nothing
    End Sub

    '��͵���������еķ���������ʽ���ַ�������ֵ�����滻/����
    '���������Զ����������ʽ,��ֵ

    Public Sub replaceByReg(patrn, Str)
        If isGet_ = false Then Call steal()
        Dim tempReg
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg = patrn
        value_ = tempReg.Replace(value_ , Str)
        Set tempReg = Nothing
    End Sub

    'Ӧ��������ʽ�Է������������ݽ��зֿ�ɼ������,��������Ϊ��<!--lkstar-->���ϵĴ��ı�/����
    'ͨ������value�õ������ݺ��������split(value,"<!--lkstar-->")�õ�����Ҫ������
    '���������Զ����������ʽ

    Public Sub pickByReg(patrn)
        If isGet_ = false Then Call steal()
        Dim tempReg, match, matches, content
        Set tempReg = New RegExp
        tempReg.IgnoreCase = true
        tempReg.Global = true
        tempReg = patrn
        Set matches = tempReg.Execute(value_)
        For Each match in matches
            content = content&match.Value&"<!--lkstar-->"
        Next
        value_ = content
        Set matches = Nothing
        Set tempReg = Nothing
    End Sub

    '���Ŵ�ģʽ���������ͷ�֮ǰӦ�ô˷���������ʱ�鿴��ػ������HTML�����ҳ����ʾЧ��/����

    Public Sub debug()
        Dim tempstr
        tempstr = "<SCRIPT>function runEx(){var winEx2 = window.open("""", ""winEx2"", ""width=500,height=300,status=yes,menubar=no,scrollbars=yes,resizable=yes""); winEx2.document.open(""text/html"", ""replace""); winEx2.document.write(unescape(event.srcElement.parentElement.children[0].value)); winEx2.document.close(); }function saveFile(){var win=window.open('','','top=10000,left=10000');win.document.write(document.all.asdf.innerText);win.document.execCommand('SaveAs','','javascript.htm');win.close();}</SCRIPT><center><TEXTAREA id=asdf name=textfield rows=32  wrap=VIRTUAL cols=""120"">"&value_&"</TEXTAREA><BR><BR><INPUT name=Button onclick=runEx() type=button value=""�鿴Ч��"">&nbsp;&nbsp;<INPUT name=Button onclick=asdf.select() type=button value=""ȫѡ"">&nbsp;&nbsp;<INPUT name=Button onclick=""asdf.value=''"" type=button value=""���"">&nbsp;&nbsp;<INPUT onclick=saveFile(); type=button value=""�������""></center>"
        'response.Write(tempstr)
		document.Write tempstr
    End Sub

End Class


