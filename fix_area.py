import sqlite3

def update_db():
    conn = sqlite3.connect('roboclass.db')
    cursor = conn.cursor()
    
    # Update old area_trabajo to the new one
    cursor.execute("""
        UPDATE usuario 
        SET area_trabajo = 'Coordinador / Administrativo' 
        WHERE area_trabajo = 'Equipo Directivo / Administrativo'
    """)
    
    conn.commit()
    print("Database updated: Old area_trabajo 'Equipo Directivo / Administrativo' changed to 'Coordinador / Administrativo'.")
    conn.close()

if __name__ == "__main__":
    update_db()
