import sqlite3

def crear_tabla():
    conn = sqlite3.connect('roboclass.db')
    cursor = conn.cursor()

    # Create the table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS configuracion_institucional (
            id INTEGER PRIMARY KEY,
            nombre_escuela VARCHAR(150),
            director VARCHAR(100),
            telefono_director VARCHAR(50),
            correo_director VARCHAR(100),
            codigo_estadistico VARCHAR(50),
            codigo_dea VARCHAR(50),
            codigo_administrativo VARCHAR(50),
            codigo_dependencia VARCHAR(50),
            codigo_sunagro VARCHAR(50),
            rif_escuela VARCHAR(50),
            rif_consejo VARCHAR(50),
            dependencia VARCHAR(50),
            ubicacion_geografica VARCHAR(50),
            clase_plantel VARCHAR(50),
            ano_fundacion VARCHAR(10),
            telefono_escuela VARCHAR(50),
            correo_escuela VARCHAR(100),
            supervisora VARCHAR(100),
            direccion TEXT,
            circuito VARCHAR(50)
        )
    """)

    # Check if a row already exists
    cursor.execute("SELECT COUNT(*) FROM configuracion_institucional")
    if cursor.fetchone()[0] == 0:
        cursor.execute("""
            INSERT INTO configuracion_institucional (
                nombre_escuela, director, telefono_director, correo_director,
                codigo_estadistico, codigo_dea, codigo_administrativo, codigo_dependencia,
                codigo_sunagro, rif_escuela, rif_consejo, dependencia,
                ubicacion_geografica, clase_plantel, ano_fundacion, telefono_escuela,
                correo_escuela, supervisora, direccion, circuito
            ) VALUES (
                'E.E.P. DANIEL O''LEARY', 'Prof. Daissy Reyna', '0424-3254033', 'daissyreyna1428@gmail.com',
                '041902', 'OD12090407', '006733370', '006733370',
                '932084', 'J303108423', 'J506628210', 'NACIONAL',
                'URBANO', 'GRADUADO', '1958', '0247-3415058',
                'eepdanieloleary@gmail.com', 'YOSEIDYS LIRA', 'AV. CARABOBO, SECTOR CASA DE ZINC', 'APU0701010ZEA'
            )
        """)
        print("Fila inicial insertada.")
    else:
        print("La tabla ya tiene datos.")

    conn.commit()
    conn.close()
    print("Tabla configuracion_institucional verificada/creada exitosamente.")

if __name__ == '__main__':
    crear_tabla()
