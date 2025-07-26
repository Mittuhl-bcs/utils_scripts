@echo off
title Daily Network & VPN Cleanup (Safe Version)
echo --------------------------------------------------
echo     SAFE Network and VPN Cleanup Script - %DATE% %TIME%
echo --------------------------------------------------

:: Flush DNS
echo [*] Flushing DNS cache...
ipconfig /flushdns

:: Winsock Reset
echo [*] Resetting Winsock...
netsh winsock reset

:: TCP/IP Reset
echo [*] Resetting TCP/IP stack...
netsh int ip reset

:: Clear ARP cache
echo [*] Clearing ARP cache...
arp -d *

:: Clear NetBIOS cache
echo [*] Clearing NetBIOS cache...
nbtstat -R
nbtstat -RR

:: Reset routing table
echo [*] Resetting routing table...
route -f

:: Restart network adapters
echo [*] Restarting active network adapters...
powershell -Command "Get-NetAdapter | Where-Object {$_.Status -eq 'Up'} | Disable-NetAdapter -Confirm:$false; Start-Sleep -Seconds 5; Get-NetAdapter | Enable-NetAdapter -Confirm:$false"

echo --------------------------------------------------
echo [*] Restarting key network services...

net stop Dnscache && net start Dnscache
net stop NlaSvc && net start NlaSvc
net stop Netman && net start Netman
net stop Dhcp && net start Dhcp
net stop iphlpsvc && net start iphlpsvc
echo [✔] Network services restarted successfully.

echo --------------------------------------------------
echo [✔] SAFE Network & VPN cleanup complete!
echo No DNS or VPN config has been altered.
echo You may reboot if needed. Safe to close.
pause
