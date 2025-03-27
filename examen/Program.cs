using System;
using System.Data;
using System.IO;
using ExcelDataReader;
using SqlClient = Microsoft.Data.SqlClient;

namespace examen
{
    class Program
    {
        static string excelFilePath = @"C:\Users\enest\OneDrive\Documentos\Examen\examen\Departamentos_Empleados.xlsx";
        static string connectionString = "Server=ALEXANDER\\SQLEXPRESS;Database=examen_adb;Trusted_Connection=True;Encrypt=False;";

        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            bool continuar = true;
            while (continuar)
            {
                Console.WriteLine("\n=== MENÚ PRINCIPAL ===");
                Console.WriteLine("1) Importar datos desde Excel");
                Console.WriteLine("2) Crear llave foránea");
                Console.WriteLine("3) Salir");
                Console.Write("Seleccione una opción: ");
                string opcion = Console.ReadLine();

                switch (opcion)
                {
                    case "1":
                        ImportarDatos();
                        break;
                    case "2":
                        CrearForeignKey();
                        break;
                    case "3":
                        continuar = false;
                        break;
                    default:
                        Console.WriteLine("Opción no válida.");
                        break;
                }
            }
            Console.WriteLine("Fin del programa.");
        }

        static void ImportarDatos()
        {
            try
            {
                // Antes de importar datos, nos aseguramos de que las tablas existan
                CrearTablas();

                DataSet ds = ReadExcelFile(excelFilePath);
                if (ds.Tables.Count > 0)
                {
                    DataTable dtDepartamentos = ds.Tables[0];
                    InsertDepartamentos(dtDepartamentos);
                }
                else
                {
                    Console.WriteLine("No se encontró la hoja de Departamentos.");
                }
                if (ds.Tables.Count > 1)
                {
                    DataTable dtEmpleados = ds.Tables[1];
                    InsertEmpleados(dtEmpleados);
                }
                else
                {
                    Console.WriteLine("No se encontró la hoja de Empleados.");
                }
                Console.WriteLine("Datos importados exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al importar datos: " + ex.Message);
            }
        }

        static void CrearForeignKey()
        {
            try
            {
                // Aseguramos que las tablas existan antes de crear la FK
                CrearTablas();

                string alterTableQuery = @"
                    IF NOT EXISTS (
                        SELECT * FROM sys.foreign_keys 
                        WHERE name = 'FK_Empleado_Departamento'
                    )
                    BEGIN
                        ALTER TABLE Empleado
                        ADD CONSTRAINT FK_Empleado_Departamento
                        FOREIGN KEY (id_departamento)
                        REFERENCES Departamento(id_departamento);
                    END
                ";

                using (var conn = new SqlClient.SqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new SqlClient.SqlCommand(alterTableQuery, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                Console.WriteLine("Llave foránea creada correctamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al crear la llave foránea: " + ex.Message);
            }
        }

        static void CrearTablas()
        {
            string createTablesScript = @"
                IF NOT EXISTS (
                    SELECT * FROM sys.objects 
                    WHERE object_id = OBJECT_ID(N'[dbo].[Departamento]') 
                      AND type in (N'U')
                )
                BEGIN
                    CREATE TABLE dbo.Departamento (
                        id_departamento INT PRIMARY KEY IDENTITY(1,1),
                        nombre VARCHAR(100) NOT NULL,
                        ubicacion VARCHAR(100) NOT NULL,
                        presupuesto INT NOT NULL,
                        fecha_creacion DATE NOT NULL DEFAULT GETDATE()
                    );
                END;

                IF NOT EXISTS (
                    SELECT * FROM sys.objects 
                    WHERE object_id = OBJECT_ID(N'[dbo].[Empleado]') 
                      AND type in (N'U')
                )
                BEGIN
                    CREATE TABLE dbo.Empleado (
                        id_empleado INT PRIMARY KEY IDENTITY(1,1),
                        nombre VARCHAR(100) NOT NULL,
                        apellido VARCHAR(100) NOT NULL,
                        salario DECIMAL(10,2) NOT NULL,
                        puesto VARCHAR(50) NOT NULL,
                        fecha_contrato DATE NOT NULL,
                        id_departamento INT NOT NULL
                    );
                END;
            ";

            using (var conn = new SqlClient.SqlConnection(connectionString))
            {
                conn.Open();
                using (var cmd = new SqlClient.SqlCommand(createTablesScript, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        static DataSet ReadExcelFile(string path)
        {
            DataSet result = null;
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var conf = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    };
                    result = reader.AsDataSet(conf);
                }
            }
            return result;
        }

        static void InsertDepartamentos(DataTable dt)
        {
            using (var conn = new SqlClient.SqlConnection(connectionString))
            {
                conn.Open();
                foreach (DataRow row in dt.Rows)
                {
                    string nombre = row["nombre"].ToString();
                    string ubicacion = row["ubicacion"].ToString();
                    string presupuestoStr = row["presupuesto"].ToString();
                    int presupuesto = 0;
                    int.TryParse(presupuestoStr, out presupuesto);

                    DateTime fechaCreacion;
                    if (!DateTime.TryParse(row["fecha_creacion"].ToString(), out fechaCreacion))
                        fechaCreacion = DateTime.Now;

                    string query = @"
                        INSERT INTO Departamento (nombre, ubicacion, presupuesto, fecha_creacion)
                        VALUES (@nombre, @ubicacion, @presupuesto, @fecha_creacion)";
                    using (var cmd = new SqlClient.SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@nombre", nombre);
                        cmd.Parameters.AddWithValue("@ubicacion", ubicacion);
                        cmd.Parameters.AddWithValue("@presupuesto", presupuesto);
                        cmd.Parameters.AddWithValue("@fecha_creacion", fechaCreacion);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        static void InsertEmpleados(DataTable dt)
        {
            using (var conn = new SqlClient.SqlConnection(connectionString))
            {
                conn.Open();
                foreach (DataRow row in dt.Rows)
                {
                    string nombre = row["nombre"].ToString();
                    string apellido = row["apellido"].ToString();
                    decimal salario = Convert.ToDecimal(row["salario"]);
                    string puesto = row["puesto"].ToString();
                    DateTime fechaContrato = DateTime.Parse(row["fecha_contrato"].ToString());
                    int id_departamento = Convert.ToInt32(row["id_departamento"]);

                    string query = @"
                        INSERT INTO Empleado (nombre, apellido, salario, puesto, fecha_contrato, id_departamento)
                        VALUES (@nombre, @apellido, @salario, @puesto, @fecha_contrato, @id_departamento)";
                    using (var cmd = new SqlClient.SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@nombre", nombre);
                        cmd.Parameters.AddWithValue("@apellido", apellido);
                        cmd.Parameters.AddWithValue("@salario", salario);
                        cmd.Parameters.AddWithValue("@puesto", puesto);
                        cmd.Parameters.AddWithValue("@fecha_contrato", fechaContrato);
                        cmd.Parameters.AddWithValue("@id_departamento", id_departamento);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
    }
}
