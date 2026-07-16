"""Definiciones ISAK Nivel 1 y columnas RIPS/UCAD."""

ISAK_VARIABLES = [
    ("MASA", "Masa corporal", "kg", "basica", 0.01, True),
    ("TALLA", "Talla / estatura", "cm", "basica", 0.01, True),
    ("PL_TRICEPS", "Pliegue tríceps", "mm", "pliegue", 0.05, True),
    ("PL_SUBESCAPULAR", "Pliegue subescapular", "mm", "pliegue", 0.05, True),
    ("PL_BICEPS", "Pliegue bíceps", "mm", "pliegue", 0.05, False),
    ("PL_CRESTA_ILIACA", "Pliegue cresta ilíaca", "mm", "pliegue", 0.05, True),
    ("PL_SUPRAESPINAL", "Pliegue supraespinal", "mm", "pliegue", 0.05, True),
    ("PL_ABDOMINAL", "Pliegue abdominal", "mm", "pliegue", 0.05, True),
    ("PL_MUSLO_ANT", "Pliegue muslo anterior", "mm", "pliegue", 0.05, True),
    ("PL_PIERNA_MED", "Pliegue pierna medial", "mm", "pliegue", 0.05, True),
    ("PR_BRAZO_REL", "Perímetro brazo relajado", "cm", "perimetro", 0.01, True),
    ("PR_BRAZO_FLEX", "Perímetro brazo flexionado", "cm", "perimetro", 0.01, False),
    ("PR_CINTURA", "Perímetro cintura (mínima)", "cm", "perimetro", 0.01, True),
    ("PR_CADERA", "Perímetro cadera", "cm", "perimetro", 0.01, True),
    ("PR_PANTORRILLA", "Perímetro pantorrilla máx.", "cm", "perimetro", 0.01, True),
    ("D_HUMERO", "Diámetro biepicondilar húmero", "cm", "diametro", 0.01, False),
    ("D_FEMUR", "Diámetro biepicondilar fémur", "cm", "diametro", 0.01, False),
]

DEPORTISTA_COLUMNS = [
    ("ID_Deportista", "ID interno"),
    ("tipoDocumentoIdentificacion", "Tipo documento RIPS"),
    ("numDocumentoIdentificacion", "Número documento"),
    ("primerApellido", "Primer apellido"),
    ("segundoApellido", "Segundo apellido"),
    ("nombres", "Nombres"),
    ("tipoUsuario", "Tipo usuario RIPS"),
    ("fechaNacimiento", "Fecha nacimiento"),
    ("codSexo", "Sexo M/F"),
    ("codPaisResidencia", "País residencia"),
    ("codMunicipioResidencia", "Municipio residencia"),
    ("codZonaTerritorialResidencia", "Zona urbana/rural"),
    ("incapacidad", "Incapacidad SI/NO"),
    ("codPaisOrigen", "País origen"),
    ("consecutivo", "Consecutivo RIPS"),
    ("telefono", "Teléfono"),
    ("email", "Correo"),
    ("eps", "EPS"),
    ("direccion", "Dirección"),
    ("deporte", "Deporte"),
    ("equipo", "Equipo"),
    ("categoria", "Categoría"),
    ("fechaIngreso", "Fecha ingreso"),
    ("activo", "Activo Sí/No"),
]

SAMPLE_DEPORTISTAS = [
    ("DEP-001", "CC", "1020304050", "García", "López", "Andrés", "04", "2001-03-15", "M", "170", "11001", "01", "NO", "170", 1, "3001234567", "andres@email.com", "Sura", "Bogotá", "Fútbol", "UCAD A", "Sub-20", "2024-01-10", "Sí"),
    ("DEP-002", "CC", "1030456789", "Martínez", "Ruiz", "María", "04", "2002-07-22", "F", "170", "11001", "01", "NO", "170", 2, "3019876543", "maria@email.com", "Sanitas", "Bogotá", "Atletismo", "UCAD B", "Senior", "2024-02-01", "Sí"),
    ("DEP-003", "TI", "1040567890", "Rodríguez", "", "Carlos", "04", "2008-11-05", "M", "170", "08001", "02", "NO", "170", 3, "3021112233", "", "Nueva EPS", "Barranquilla", "Natación", "UCAD A", "Juvenil", "2024-03-15", "Sí"),
    ("DEP-004", "CC", "1050678901", "Hernández", "Vega", "Laura", "04", "2000-01-30", "F", "170", "11001", "01", "NO", "170", 4, "3104445566", "laura@email.com", "Compensar", "Bogotá", "Baloncesto", "UCAD C", "Senior", "2023-11-20", "Sí"),
    ("DEP-005", "CC", "1060789012", "Acosta", "Pérez", "Diego", "04", "1999-09-12", "M", "170", "05001", "01", "NO", "170", 5, "3207778899", "diego@email.com", "Sura", "Medellín", "Ciclismo", "UCAD B", "Elite", "2023-08-05", "Sí"),
]

# Evaluaciones de muestra: (id, fecha, masa, talla, pl_triceps, pl_subesc, pl_biceps, pl_cresta, pl_supra, pl_abd, pl_muslo, pl_pierna, pr_brazo, pr_flex, pr_cintura, pr_cadera, pr_pant, d_hum, d_fem, antropometrista, bloqueado)
SAMPLE_SEGUIMIENTO = [
    ("DEP-001", "2025-10-01", 78.2, 178, 8.5, 12.1, 5.2, 10.3, 9.8, 15.2, 18.5, 12.4, 28.5, 32.1, 78.0, 95.2, 36.8, 6.8, 9.5, "Eder Acosta", "Sí"),
    ("DEP-001", "2025-11-15", 77.1, 178, 8.0, 11.5, 5.0, 9.8, 9.2, 14.5, 17.8, 11.9, 28.2, 31.8, 76.5, 94.5, 36.5, 6.8, 9.5, "Eder Acosta", "Sí"),
    ("DEP-001", "2026-01-10", 76.5, 178, 7.8, 11.0, 4.8, 9.5, 8.9, 14.0, 17.2, 11.5, 28.0, 31.5, 75.8, 94.0, 36.2, 6.9, 9.6, "Eder Acosta", "Sí"),
    ("DEP-002", "2025-10-05", 62.5, 168, 14.2, 16.5, 8.1, 14.8, 13.5, 18.2, 22.5, 16.8, 26.5, 29.8, 68.5, 98.5, 35.2, 5.8, 8.2, "Eder Acosta", "Sí"),
    ("DEP-002", "2025-12-20", 61.8, 168, 13.8, 16.0, 7.9, 14.2, 13.0, 17.5, 21.8, 16.2, 26.2, 29.5, 67.8, 97.8, 34.8, 5.8, 8.2, "Eder Acosta", "Sí"),
    ("DEP-002", "2026-02-01", 61.2, 168, 13.5, 15.5, 7.7, 13.8, 12.6, 17.0, 21.2, 15.8, 26.0, 29.2, 67.2, 97.2, 34.5, 5.9, 8.3, "Eder Acosta", "Sí"),
    ("DEP-003", "2025-11-01", 55.0, 172, 6.5, 9.8, 4.2, 7.5, 7.2, 11.5, 14.8, 9.5, 25.5, 28.5, 72.0, 88.5, 34.2, 6.2, 8.8, "Eder Acosta", "Sí"),
    ("DEP-003", "2026-01-20", 56.2, 173, 6.8, 10.0, 4.3, 7.8, 7.5, 11.8, 15.0, 9.8, 26.0, 29.0, 71.5, 89.0, 34.5, 6.3, 8.9, "Eder Acosta", "Sí"),
    ("DEP-004", "2025-10-10", 68.5, 175, 12.5, 14.8, 7.5, 12.5, 11.8, 16.5, 20.5, 14.5, 27.8, 31.2, 72.5, 96.5, 35.8, 6.5, 9.0, "Eder Acosta", "Sí"),
    ("DEP-004", "2026-01-05", 67.8, 175, 12.0, 14.2, 7.2, 12.0, 11.2, 16.0, 19.8, 14.0, 27.5, 30.8, 71.8, 95.8, 35.5, 6.5, 9.0, "Eder Acosta", "Sí"),
    ("DEP-005", "2025-09-15", 72.0, 180, 7.2, 10.5, 4.8, 8.5, 8.0, 12.8, 16.5, 10.5, 29.5, 33.5, 76.5, 92.5, 37.5, 7.2, 10.0, "Eder Acosta", "Sí"),
    ("DEP-005", "2025-12-01", 71.2, 180, 6.8, 10.0, 4.5, 8.0, 7.5, 12.2, 16.0, 10.0, 29.2, 33.0, 75.8, 92.0, 37.2, 7.2, 10.0, "Eder Acosta", "Sí"),
    ("DEP-005", "2026-02-15", 70.5, 180, 6.5, 9.5, 4.3, 7.8, 7.2, 11.8, 15.5, 9.8, 29.0, 32.8, 75.2, 91.5, 37.0, 7.3, 10.1, "Eder Acosta", "Sí"),
]

ADMIN_PASSWORD = "UCAD2026"
TIPO_DOC_LIST = "CC,TI,CE,RC,PA,MS,AS"
TIPO_USUARIO_LIST = "01,02,03,04"
SEXO_LIST = "M,F"
ZONA_LIST = "01,02"
SI_NO_LIST = "SI,NO,Sí,No"
ACTIVO_LIST = "Sí,No"
LADO_LIST = "D,I"
METODO_LIST = "A,B"
