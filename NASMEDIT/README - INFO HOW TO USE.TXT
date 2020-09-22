Scroll down for HOWTO tutorial.

I made this program becouse ASM is so hard to learn, and I simply don't want to start learning it if I can't
do much. When using ASM in VB I can learn ASM much quicker. Because my program supports commenting
it is much easyer to learn ASM. When I began with this project, I knew nothing about ASM. And now I've made
the example ASM (the one wich is shown in the VB example program).

You can use this program for free, and you can use modAsm in every project you want to use it in!


How to use NasmEditor:
1. Download nasm from
   http://sourceforge.net/project/showfiles.php?group_id=6208
   (Download the Win32 binaries)

2. Unzip or unrar it.
3. Open the NasmEditor project
4. Compile the NasmEditor project
5. Copy the NasmEditor EXE to the folder containing nasmw.exe

How to use NasmScripts in VB:
1. Open nasmEditor
2. Start an new Multiply function example (or an callwindowproc project if you understand asm)
3. Go to build->Test if compatible with VB, to test if it's working
4. Go to build->Translate to hexdata, and copy the bottom text.
5. Now you have something like Const MyConst = "3A002B..."
6. Paste that const into your new project (in your form)
7. Add modasm.bas to your project
8. if you want to call your asm function:
   [CODE]
      asmInject MyConst       'To inject the asm data into your program
      MsgBox asmRun(P1, P2, P3, P4)
   [CODE]
-If you still don't know how check the "HOW TO USE ASM IN VB" example!-