Set debug = new MyDebug

debug.wsh_info()
Set T = new Targets
T.create("test")
T.hit("test")

debug.print("array test")

Dim Ts 
Ts = Array(New Target, New Target, New Target)
Ts(0).name = "t1"
Ts(1).name = "t2"
Ts(2).name = "t3"
Ts(0).Hit() : Ts(1).Hit():Ts(2).Hit()

debug.print("end of script")
