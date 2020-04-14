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
Trabajaba en la compañia de seguros Cruz del Sur en el departamento informatica migrando el modulo de reaseguros (que wea de negocio mas enredada) por alla por el año 2000 y la informacion estaba almacenada en un sistema AS/400 base de datos DB2. Por ese entonces la unica forma de consultar datos o ejecutar actualizaciones de las tablas era a traves de una interfaz de "fosforo de verde" de una terminal. 

Como no habia una forma "amigable" y simple de conectarse a la base de datos DB/2 fue que nacio inicialmente un proyecto llamado "SymphonyX" en honor al grupo de powermetal americano. Primero inicialmente solo era una interfaz simple para conectarse y tener una ventana unica con el resultado de la ejecucion.

Posteriormente y a sugerencia de mis compañeros de trabajo de la epoca fue que se cambio a multidocumento (MDI) y la posibilidad ademas de poder conectarse a otras fuentes de datos usando en ese tiempo el protocolo ODBC.

Para esa epoca la herramienta curiosamente llamo la atencion de toda la gerencia TI de ese entonces y se convirtio en la herramienta "oficial" de consulta a la base de datos DB/2.

# Archivos Necesarios
Este proyecto ocupa 5 componentes ActiveX 

- Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#..\..\..\..\..\..\..\Windows\SysWOW64\stdole2.tlb#OLE Automation
- Reference=*\G{00000205-0000-0010-8000-00AA006D2EA4}#2.5#0#..\..\..\..\..\..\..\Program Files (x86)\Common Files\System\ado\msado25.tlb#Microsoft ActiveX Data Objects 2.5 Library
- Reference=*\G{00000600-0000-0010-8000-00AA006D2EA4}#2.5#0#..\..\..\..\..\..\..\Program Files (x86)\Common Files\System\ado\msadox.dll#Microsoft ADO Ext. 2.5 for DDL and Security
- Object={3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0; RICHTX32.OCX
- Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX
- Object={F13C99C0-4D73-11D2-B8B2-0000C00A958C}#3.0#0; fpSpr30.ocx

El archivo PVB_XMENU.DLL es un componente customizado para que los menus se puedan aplicar iconos y ayuda al momento de selección.

* NOTA: Este proyecto ocupaba un componente activex de pago llamado Farpoint Spread el cual se usaba la version 3.0.

# Registro de los componentes ActiveX
Se debe realizar desde la linea de comando de windows regsvr32.exe [nombre del componente]
Para windows 10 necesitaras instalar con permisos de administrador. 

# Notas de los componentes ActiveX de Windows
Si obtienes error de licencia de componentes al momento de ejecutar el proyecto necesitaras instalar quizas la runtime de Visual Basic 5 (MSCVBM50.DLL) y bajar el archivo VB5CLI.EXE y VBUSC.EXE ambos disponibles en internet para descarga. Esto corregira los problemas de licencia de componentes de VB5.

# Desarrollo del proyecto
El proyecto se desarrollo usando la interfaz multiple de documentos MDI, la libreria de conexion a base de datos DAO y inicialmente el componente listview para el resultado de datos. Para el manejo de las multiples conexiones se creo un arreglo el cual almacenaba segun el indice la conexion a la base de datos seleccionada.

Cada ventana almacena de forma interna a cual conexion pertenece y con eso se diferencia el como se conecta al origen de datos.

Para colorear los comandos de las sentencias SQL se ocupa el componente active RTF.

Como nota curiosa de este proyecto no faltaron los compañeros de trabajo que tambien se animaron a desarrollar algo similar incluso no falto el care raja que se atribuyo el proyecto como su autoria y aquel que tomo las buenas ideas de mi proyecto y le coloco de nombre "natacha" ...

# Freeware
Por esos años mi intención fue ofrecerlo gratis a la comunidad Visual Basic que era bastante activa por esos años. Para esto levante un sitio web donde tenia varias otras aplicaciones que tambien habian sido creadas de la necesidad y que las distribuia de forma gratis.

# Palabras Finales
Espero que este proyecto que nacio de una necesidad personal sea usado con motivos de estudio y motivación. De como se pueden copiar las buenas ideas y mejorarlas. 

