To Use:

Requires word 97 or greater 
Place the ocx in system directory
I am not sure what languages it works for: English only tested

Other:

If you want source for ocx e-mail me roboman1@email.com
If you use this please include me in an about screen if you have one!

functions:

-------------------------------------------------------------------------------------

wordlink - starts/ends the connection to Microsoft Word

ex.) Required for spelling check! If not you will get an automation error!

Private Sub Form_Load()
scheck1.wordlink
End Sub

-------------------------------------------------------------------------------------

showmessages(showit as boolean) - tells wether the control should show Messageboxes

ex.) 

scheck1.showmessages(true) 'Shows the Message boxes

-------------------------------------------------------------------------------------

checkspell(checktext as string) as string - checks spelling
returns correct spelling of words

ex.)

text1.text = scheck1.checkspell(text1.text)

-------------------------------------------------------------------------------------

countwords(checktext as string) as integer - counts word used in string
returns word count - small problems with contractions - counts a contraction as 2 words

ex.)

lblwords.caption = scheck1.countwords(text1.text)

-------------------------------------------------------------------------------------

countchar(checktext as string) as integer - counts characters
returns character count with no problems!

ex.)

lblchar.caption = scheck1.countchar(text1.text)