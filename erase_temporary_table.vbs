Const wdAlignParagraphCenter = 1
Const wdCollapseEnd = 0
Const wdstory = 6
Const wdAlignRowCenter = 1
Const fig_name_common = "�}*:"
Const table_name_common = "�\*:"


' �J�����g�t�H���_�̎擾
Dim shell_obj
Dim current_dir
Set shell_obj = CreateObject( "WScript.Shell" )
current_dir = shell_obj.CurrentDirectory
Set shell_obj = Nothing

' ���C������
main(current_dir)

Sub main(current_dir)
    Set word_obj = CreateObject("Word.Application")
    word_obj.Visible = True

    target_filename = current_dir & "\trial.docx"

    Set doc_obj = word_obj.Documents.Open(target_filename)
    
    Call erase_temporary_table(word_obj,doc_obj, "[erase]")

End Sub

Sub erase_temporary_table(word_obj,doc_obj, keyword)
    For Each table in doc_obj.Tables
        table.Delete
    Next
End Sub


