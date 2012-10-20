echo off
mktyplib /tlb SimplyVBUnitType.tlb SimplyVBUnitType.odl
if not errorlevel 1 goto end
pause
:end
regtlib SimplyVBUnitType.tlb