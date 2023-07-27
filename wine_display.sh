if [ ! -f ~/.wine ]
then
  WINEARCH=win32 winecfg
fi
if [ ! -f /tmp/.X99-lock ]
then
  Xvfb :99 &
fi
export DISPLAY=:99
