# inventario-biomedico
Sistema de gestión de inventario para ingeniería biomédica del HDLM
Bienvenido al Sistema de Gestión de Inventario Biomédico. Este proyecto fue creado para el departamento de Ingeniería Clínica del Hospital de la Mujer, con el objetivo de dejar atrás los registros manuales y tener una plataforma web rápida para controlar los equipos médicos, los mantenimientos y las bajas.

Para que entiendas rápido cómo está armado, el sistema está hecho en Python usando una herramienta llamada Streamlit para la parte visual, y los datos se guardan en la nube usando una base de datos llamada Neon.

Si revisas los archivos del proyecto, esto es lo que hace cada uno: el archivo "inventario_hospital.py" es el cerebro de todo y ahí está el código principal; el archivo "requirements.txt" es la lista de herramientas que necesitas instalar en tu computadora para que el programa corra; "issea.png" es el logo que aparece en la pantalla; y ".gitignore" es un archivo de seguridad que evita que se suban contraseñas a internet.

Si eres un nuevo practicante y te toca hacerle mejoras al código en tu computadora, hay un paso crítico que debes conocer. Por seguridad, las contraseñas de la base de datos no están guardadas aquí. Para que el sistema funcione en tu compu, tienes que crear una carpeta oculta llamada ".streamlit" y adentro hacer un archivo de texto llamado "secrets.toml". En ese archivo debes escribir las claves de acceso de Neon y la contraseña de la página (pídele estos datos al ingeniero a cargo). Si te saltas este paso, el sistema te marcará error porque no tendrá permiso para ver los datos.

Una vez que tengas tus contraseñas listas, solo abre tu terminal y escribe "streamlit run inventario_hospital.py" para ver la página funcionando. Y no te preocupes por la página oficial que ya está en internet; cada vez que guardes tus cambios terminados aquí en GitHub, la plataforma en la nube se actualizará solita. ¡Cuida mucho este sistema y éxito en tus prácticas!
