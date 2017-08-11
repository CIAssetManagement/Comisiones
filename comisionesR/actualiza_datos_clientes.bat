del "tablas/posicionMes.csv"
echo "Yes" | copy "\\192.168.0.68\fidem\UPDATE\CIESTRATEGIAS\posicion_cism.csv" "tablas/posicionMes.csv"
rem Por alguna muy rara razon de AM hay que pedir configClientes dos veces, la primera vez no sirve....
del "tablas/carterasModeloClientes.txt"
"C:/Program Files (x86)/FIDEM/AM/AmConsola.exe" "2|%AMUSER%|%PWD%|8|configClientes|C:/Github/comisiones/comisionesR/|carterasModeloClientes|"
move /Y "carterasModeloClientes.txt" "tablas\"
