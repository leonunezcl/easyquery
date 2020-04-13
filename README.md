# Programa
Easy Query - Multi Gestor de bases de datos 

# Autor
Luis Leonardo Nuñez Ibarra. Año 2000 - 2003. email : leo.nunez@gmail.com. 

Chileno, casado , tengo 2 hijos. Aficionado a los videojuegos y el tenis de mesa. Mi primer computador fue un Talent MSX que me compro mi papa por alla por el año 1985. En el di mis primeros pasos jugando juegos como Galaga y PacMan y luego programando en MSX-BASIC. 

En la actualidad mi area de conocimiento esta referida a las tecnologias .NET con mas de 15 años de experiencia desarrollando varias paginas web usando asp.net con bases de datos sql server y Oracle. Integrador de tecnologias, desarrollo de servicios, aplicaciones de escritorio.

# Tipo de Proyecto
Easy Query es un multi gestor de bases de datos. La idea detras de este proyecto era tener centralizado en una sola aplicación el poder trabajar con distintos tipos bases de datos usando conexiones via ODBC

# Prologo
Regala un pescado a un hombre y le darás alimento para un día, enseñale a pescar y lo alimentarás para el resto de su vida (Proverbio Chino)

# Historia


# Archivos Necesarios
Este proyecto ocupa 4 componentes ActiveX 

Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\SYSTEM\StdOle2.tlb#OLE Automation
Reference=*\G{8B217740-717D-11CE-AB5B-D41203C10000}#1.0#0#C:\WINDOWS\SYSTEM\TLBINF32.DLL#TypeLib Information
Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
Reference=*\G{69EDFBA5-9FEC-11D5-89A4-F0FAEF3C8033}#1.0#0#C:\WINDOWS\SYSTEM\PVB_XMENU.DLL#PVB6 ActiveX DLL - Menu With Bitmaps !

El archivo PVB_XMENU.DLL es un componente customizado para que los menus se puedan aplicar iconos y ayuda al momento de selección.

# Registro de los componentes ActiveX
Se debe realizar desde la linea de comando de windows regsvr32.exe [nombre del componente]
Para windows 10 necesitaras instalar con permisos de administrador. 
La libreria vbzip10.dll la debes copiar en el directorio WINDOWS\SYSTEM de tu computador.

# Notas de los componentes ActiveX de Windows
Si obtienes error de licencia de componentes al momento de ejecutar el proyecto necesitaras instalar quizas la runtime de Visual Basic 5 (MSCVBM50.DLL) y bajar el archivo VB5CLI.EXE y VBUSC.EXE ambos disponibles en internet para descarga. Esto corregira los problemas de licencia de componentes de VB5.

# Desarrollo del proyecto
No fue muy dificil idear el concepto y desarrollo de este utlitario ya que teniendo como base el analizador de proyectos que tenia desarrollado en Proyect Explorer el desafio era poder incorporar la rutina de compresión .zip. Para esto y como en la mayoria de mis proyectos personales revise si steve de vbaccelerator ya tenia algo con .zip desarrollado.

Para mi buena fortuna ya tenia varios prototipos y clases desarrolladas para usar para comprimir y descomprimir archivos .zip. Con esto ya solucionado solo fue cosa de tiempo desarrollar una interfaz basica que tuviera todo lo que necesitaba para poder respaldar mis proyectos y todos los archivos relacionados de forma centralizada en un unico archivo .zip.

# Freeware
Por esos años mi intención fue ofrecerlo gratis a la comunidad Visual Basic que era bastante activa por esos años. Para esto levante un sitio web donde tenia varias otras aplicaciones que tambien habian sido creadas de la necesidad y que las distribuia de forma gratis.

# Palabras Finales
Espero que este proyecto que nacio de una necesidad personal sea usado con motivos de estudio y motivación. De como se pueden copiar las buenas ideas y mejorarlas. 

