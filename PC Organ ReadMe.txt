This program is quite basic as it is meant to demonstrate the 'Beep' API call in 'Kernel32.Dll'

This is not a complex API Call, the only arguments necessary are 'dwFreq' and 'dwDuration'.

'dwFreq' (Long) specifies the note to play by frequency (eg. 440 = A)
'dwDuration' (Long) specifies the duration of the note in miliseconds (eg. 1000 = 1 Sec.) 



I went a little overboard with mapping each note, but I was just playing around with it.
However, it turned out good, I was able to use the Control Arrays for playback by control 
position (Top) and Color (Fillcolor) to determine note frequency. It even has a pause note 
feature otherwise it plays through all the notes without rest.

Here is an exaple: Play on keyboard in sequence shown here and press playback with Tempo at 100

Then reapeat without pause notes to hear the reason I added them ;)

P = Pause Note




____P_______P________P________P__________________________
_________________________________________________________
_________________________________________________________
__________A______________________________________________
_________________________________________________________
______G_______G__________________________________________
_________________________________________________________
________________F_________F______________________________
_________________________________________________________
__________________D#___D#_________D#_____________________
____________________________D___D________________________
_________________________________________________________
__C_____C____________________________C___________________   

Next Version will allow you to save and open PCO Files, copy/cut & paste Sections, 
and edit individual notes (hopefully...).

Enjoy!

killer_cobra@hotmail.com