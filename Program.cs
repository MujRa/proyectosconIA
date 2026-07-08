using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OutlookPdfSorter
{
    // ─────────────────────────────────────────────────────────────────────────
    // MODELOS DE CONFIGURACIÓN
    // ─────────────────────────────────────────────────────────────────────────

    public class MySQLConfig
    {
        public string Server   { get; set; } = "localhost";
        public string Database { get; set; } = "overweb";
        public string User     { get; set; } = "root";
        public string Password { get; set; } = "";
        public int    Port     { get; set; } = 3306;

        public string ConnectionString =>
            $"Server={Server};Port={Port};Database={Database};" +
            $"Uid={User};Pwd={Password};CharSet=utf8;SslMode=None;";
    }

    public class AppConfig
    {
        public MySQLConfig MySQL               { get; set; } = new();
        public string      RutaBaseRed         { get; set; } = "";
        public string      CarpetaTemporal     { get; set; } = "C:\\PDFs\\_temp";
        public string      CarpetaErrores      { get; set; } = "C:\\PDFs\\_no_clasificados";
        public string      NombreCuentaOutlook { get; set; } = "";
        public bool        ModoDebug           { get; set; } = true;
        public string      CarpetaLog          { get; set; } = "C:\\PDFs\\logs";
        // Credenciales para \\serveroversea\oversea2026 (26PC y 26TR)
        public string      ServidorOversea2026 { get; set; } = "\\\\serveroversea\\oversea2026";
        public string      UsuarioOversea2026  { get; set; } = "RafaelM";
        public string      PasswordOversea2026 { get; set; } = "4ndr0m3d4*";
    }

    public class ResultadoPdf
    {
        public string  NombreArchivo     { get; set; } = "";
        public string? CodigoEncontrado  { get; set; }
        public string? FacturaEncontrada { get; set; }
        public string  FuenteCodigo      { get; set; } = "";
        public string? NombreCliente     { get; set; }
        public string? RutaDestino       { get; set; }
        public bool    Exitoso           { get; set; }
        public string  Error             { get; set; } = "";
    }

    // ─────────────────────────────────────────────────────────────────────────
    // PROGRAMA PRINCIPAL
    // ─────────────────────────────────────────────────────────────────────────

    internal class Program
    {
        private static readonly Regex CodigoRegex = new(
            @"\b(26(?:PC|PV|TR)\d+)(?:/\d{4})?\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static AppConfig   _cfg       = new();
        private static StreamWriter? _logWriter = null;
        private static string      _rutaLog   = "";
        private static HashSet<string> _correosYaProcesados = new();
        private static string      _rutaRegistroProcesados = "";

        // ── P/Invoke: GetActiveObject (reemplaza Marshal.GetActiveObject en .NET 5+) ──
        [DllImport("ole32.dll")]
        private static extern int GetActiveObject(
            ref Guid rclsid,
            IntPtr   pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        // ── P/Invoke: WNetAddConnection2 para conectar recurso de red con credenciales ──
        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        private static extern int WNetAddConnection2(
            ref NETRESOURCE netResource,
            string          password,
            string          username,
            uint            dwFlags);

        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        private static extern int WNetCancelConnection2(
            string lpName,
            uint   dwFlags,
            bool   fForce);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct NETRESOURCE
        {
            public uint   dwScope;
            public uint   dwType;       // RESOURCETYPE_DISK = 1
            public uint   dwDisplayType;
            public uint   dwUsage;
            public string lpLocalName;
            public string lpRemoteName;
            public string lpComment;
            public string lpProvider;
        }

        // ── ENTRADA ──────────────────────────────────────────────────────────
        static void Main(string[] args)
        {
            IniciarLog();
            try
            {
                Titulo("Clasificador de PDFs Oversea — 26PC / 26PV / 26TR");
                CargarConfig();
                PrepararCarpetas();
                CargarRegistroProcesados();
                ProbarConexionMySQL();
                ConectarRecursoDeRed();   // conecta \\serveroversea\oversea2026

                dynamic outlookApp = ObtenerOutlookApp();
                try
                {
                    var correos = ObtenerCorreosDeHoy(outlookApp);
                    Log($"Correos de hoy con adjuntos PDF: {correos.Count}");

                    int totalPdf = 0, ok = 0, sinCodigo = 0, sinCliente = 0, errores = 0;

                    foreach (dynamic correo in correos)
                    {
                        string entryId = "";
                        try { entryId = (string)correo.EntryID; } catch { }

                        // Saltar si este correo ya fue procesado en una ejecución anterior
                        if (!string.IsNullOrEmpty(entryId) && _correosYaProcesados.Contains(entryId))
                        {
                            Log($"  [YA PROCESADO] Omitiendo: {correo.Subject ?? "(sin asunto)"}");
                            continue;
                        }

                        string asunto = correo.Subject ?? "(sin asunto)";
                        Log($"\n  Correo: {asunto}");

                        bool alMenosUnPdfProcesado = false;
                        dynamic adjuntos = correo.Attachments;
                        int nAdj = adjuntos.Count;

                        for (int i = 1; i <= nAdj; i++)
                        {
                            dynamic adj      = adjuntos[i];
                            string  fileName = adj.FileName;

                            if (!fileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                                continue;

                            totalPdf++;
                            alMenosUnPdfProcesado = true;
                            var resultado = ProcesarAdjunto(adj, fileName);
                            MostrarResultado(resultado);

                            if      (resultado.Exitoso)                  ok++;
                            else if (resultado.CodigoEncontrado == null) sinCodigo++;
                            else if (resultado.NombreCliente    == null) sinCliente++;
                            else                                         errores++;
                        }

                        // Marcar el correo como procesado solo si tenía adjuntos PDF
                        if (alMenosUnPdfProcesado && !string.IsNullOrEmpty(entryId))
                        {
                            _correosYaProcesados.Add(entryId);
                            GuardarRegistroProcesados();
                        }
                    }

                    Resumen(totalPdf, ok, sinCodigo, sinCliente, errores);

                    // Limpiar carpetas temporales al finalizar
                    LimpiarCarpeta(_cfg.CarpetaTemporal);
                    LimpiarCarpeta(_cfg.CarpetaErrores);
                }
                finally
                {
                    Marshal.ReleaseComObject(outlookApp);
                }
            }
            catch (Exception ex)
            {
                MostrarError($"ERROR FATAL: {ex.Message}");
                MostrarError(ex.StackTrace ?? "");
            }

            CerrarLog();
        }

        // ── CONFIGURACIÓN ────────────────────────────────────────────────────

        static void CargarConfig()
        {
            string ruta = Path.Combine(AppContext.BaseDirectory, "config.json");
            if (!File.Exists(ruta))
                throw new FileNotFoundException($"No se encontró config.json en: {ruta}");

            _cfg = JsonSerializer.Deserialize<AppConfig>(
                       File.ReadAllText(ruta),
                       new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                   ?? throw new InvalidOperationException("config.json vacío o mal formado.");

            Log($"Config OK — BD: {_cfg.MySQL.Database}@{_cfg.MySQL.Server}");
        }

        static void PrepararCarpetas()
        {
            Directory.CreateDirectory(_cfg.CarpetaTemporal);
            Directory.CreateDirectory(_cfg.CarpetaErrores);
        }

        // ── RED — conectar \\serveroversea\oversea2026 con credenciales ───────

        static void ConectarRecursoDeRed()
        {
            try
            {
                var nr = new NETRESOURCE
                {
                    dwType      = 1,   // RESOURCETYPE_DISK
                    lpLocalName = "",  // sin letra de unidad, solo conexión
                    lpRemoteName = _cfg.ServidorOversea2026,
                    lpProvider  = ""
                };

                // Intentar desconectar primero por si había una sesión anterior
                WNetCancelConnection2(_cfg.ServidorOversea2026, 0, true);

                int resultado = WNetAddConnection2(
                    ref nr,
                    _cfg.PasswordOversea2026,
                    _cfg.UsuarioOversea2026,
                    0);

                if (resultado == 0)
                    Log($"Red OK — conectado a {_cfg.ServidorOversea2026} como {_cfg.UsuarioOversea2026}");
                else if (resultado == 1219)  // ERROR_SESSION_CREDENTIAL_CONFLICT — ya conectado
                    Log($"Red OK — ya existía sesión en {_cfg.ServidorOversea2026}");
                else
                    LogWarn($"WNetAddConnection2 devolvió código {resultado} para {_cfg.ServidorOversea2026}");
            }
            catch (Exception ex)
            {
                LogWarn($"No se pudo conectar a {_cfg.ServidorOversea2026}: {ex.Message}");
            }
        }

        // ── MYSQL ─────────────────────────────────────────────────────────────

        static void ProbarConexionMySQL()
        {
            using var conn = new MySqlConnection(_cfg.MySQL.ConnectionString);
            conn.Open();
            Log($"MySQL OK — {_cfg.MySQL.Server}/{_cfg.MySQL.Database}");
        }

        static string? ObtenerNombreCliente(string codigo)
        {
            string codigoserie = codigo.Substring(0, 4);
            string textonumero = codigo.Substring(4);
            int    numero      = int.Parse(textonumero);

            const string sql = @"
                SELECT nombrecliente
                FROM   clientespedidos
                WHERE  UPPER(TRIM(numero)) = UPPER(@numero)
                  AND  UPPER(TRIM(serie))  = UPPER(@codigoserie)
                LIMIT  1;";

            using var conn = new MySqlConnection(_cfg.MySQL.ConnectionString);
            conn.Open();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@numero",      numero);
            cmd.Parameters.AddWithValue("@codigoserie", codigoserie.ToUpper());
            return cmd.ExecuteScalar() as string;
        }

        static string? ObtenerFacturaCliente(string codigo)
        {
            string codigoserie = codigo.Substring(0, 4);
            string textonumero = codigo.Substring(4);
            int    numero      = int.Parse(textonumero);

            const string sql = @"
                SELECT CONCAT(codigofacserie, codigofacnumero) AS facturacliente
                FROM   clientespedidos
                WHERE  UPPER(TRIM(numero)) = UPPER(@numero)
                  AND  UPPER(TRIM(serie))  = UPPER(@codigoserie)
                LIMIT  1;";

            using var conn = new MySqlConnection(_cfg.MySQL.ConnectionString);
            conn.Open();
            using var cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@numero",      numero);
            cmd.Parameters.AddWithValue("@codigoserie", codigoserie.ToUpper());
            return cmd.ExecuteScalar() as string;
        }

        // ── OUTLOOK — COM LATE BINDING ────────────────────────────────────────

        static dynamic ObtenerOutlookApp()
        {
            Log("Conectando a Outlook...");
            try
            {
                Type? t = Type.GetTypeFromProgID("Outlook.Application")
                          ?? throw new Exception("Outlook no registrado en COM.");
                Guid clsid = t.GUID;
                int  hr    = GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
                if (hr == 0)
                {
                    Log("  -> Reutilizando instancia de Outlook ya abierta.");
                    return obj;
                }
            }
            catch (Exception ex) when (!ex.Message.Contains("no registrado"))
            {
                // sin instancia activa, continuamos
            }

            Type? outlookType = Type.GetTypeFromProgID("Outlook.Application")
                                ?? throw new InvalidOperationException("Outlook no instalado.");
            dynamic app = Activator.CreateInstance(outlookType)!;
            Log("  -> Outlook iniciado por el programa.");
            return app;
        }

        static List<dynamic> ObtenerCorreosDeHoy(dynamic outlookApp)
        {
            dynamic ns = outlookApp.GetNamespace("MAPI");
            dynamic inbox;

            if (!string.IsNullOrWhiteSpace(_cfg.NombreCuentaOutlook))
            {
                dynamic stores            = ns.Stores;
                dynamic? cuentaEncontrada = null;
                for (int i = 1; i <= stores.Count; i++)
                {
                    dynamic store = stores[i];
                    if (string.Equals((string)store.DisplayName,
                            _cfg.NombreCuentaOutlook, StringComparison.OrdinalIgnoreCase))
                    {
                        cuentaEncontrada = store;
                        break;
                    }
                }
                if (cuentaEncontrada == null)
                    throw new InvalidOperationException(
                        $"Cuenta '{_cfg.NombreCuentaOutlook}' no encontrada en Outlook.");

                inbox = cuentaEncontrada.GetDefaultFolder(6); // olFolderInbox = 6
            }
            else
            {
                inbox = ns.GetDefaultFolder(6);
            }

            var lista = new List<dynamic>();
            RecogerCorreosDeHoy(inbox, lista, 0);
            return lista;
        }

        /// <summary>
        /// Recorre recursivamente la carpeta y todas sus subcarpetas buscando
        /// correos de hoy que tengan adjuntos. Se registra en el log cada carpeta visitada.
        /// </summary>
        static void RecogerCorreosDeHoy(dynamic carpeta, List<dynamic> lista, int nivel)
        {
            var hoyInicio = DateTime.Today;
            var hoyFin    = hoyInicio.AddDays(1);
            // DASL requiere formato MM/dd/yyyy (en-US)
            string filtro =
                $"[ReceivedTime] >= '{hoyInicio:dd/MM/yyyy HH:mm}' " +
                $"AND [ReceivedTime] < '{hoyFin:dd/MM/yyyy HH:mm}'";

            string indentacion = new string(' ', nivel * 2);
            string nombreCarpeta = "(desconocida)";

            try
            {
                nombreCarpeta = (string)carpeta.Name;
                Log($"{indentacion}Carpeta: {nombreCarpeta}");

                // Correos de hoy en esta carpeta
                dynamic filtrado = carpeta.Items.Restrict(filtro);
                foreach (dynamic item in filtrado)
                {
                    try
                    {
                        if ((int)item.Class == 43 && (int)item.Attachments.Count > 0)
                            lista.Add(item);
                    }
                    catch { }
                }

                // Recorrer subcarpetas recursivamente
                dynamic subcarpetas = carpeta.Folders;
                int total = subcarpetas.Count;
                for (int i = 1; i <= total; i++)
                {
                    try
                    {
                        RecogerCorreosDeHoy(subcarpetas[i], lista, nivel + 1);
                    }
                    catch (Exception ex)
                    {
                        LogWarn($"{indentacion}  No se pudo acceder a subcarpeta #{i}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogWarn($"{indentacion}Error en carpeta '{nombreCarpeta}': {ex.Message}");
            }
        }

        // ── LÓGICA PRINCIPAL DE CADA ADJUNTO ─────────────────────────────────

        static ResultadoPdf ProcesarAdjunto(dynamic adj, string fileName)
        {
            var res = new ResultadoPdf { NombreArchivo = fileName };

            // PASO 1 — Buscar código en el nombre del archivo
            res.CodigoEncontrado = ExtraerCodigo(fileName);
            if (res.CodigoEncontrado != null)
            {
                Log($"      [Nombre archivo] Codigo: {res.CodigoEncontrado}");
                res.FuenteCodigo = "NombreArchivo";
            }
            else
            {
                Log($"      [Nombre archivo] Sin codigo -> abriendo PDF...");
            }

            string rutaTemp = Path.Combine(_cfg.CarpetaTemporal, $"{Guid.NewGuid():N}_{fileName}");

            try
            {
                adj.SaveAsFile(rutaTemp);

                // PASO 2 — Si no estaba en el nombre, leer el PDF
                if (res.CodigoEncontrado == null)
                {
                    res.CodigoEncontrado = BuscarCodigoEnPdf(rutaTemp);
                    if (res.CodigoEncontrado != null)
                    {
                        Log($"      [Contenido PDF] Codigo: {res.CodigoEncontrado}");
                        res.FuenteCodigo = "ContenidoPDF";
                    }
                    else
                    {
                        Log($"      [Contenido PDF] Codigo NO encontrado.");
                    }
                }

                // PASO 3 — Sin código: eliminar temporal y salir
                if (res.CodigoEncontrado == null)
                {
                    EliminarArchivo(rutaTemp);
                    res.Error = "No se encontró código 26PC/26PV/26TR";
                    return res;
                }

                // PASO 4 — Enrutar según prefijo
                string prefijo = res.CodigoEncontrado.Substring(0, 4).ToUpper(); // 26PC / 26PV / 26TR

                if (prefijo == "26PV")
                {
                    // ── 26PV: consulta MySQL → carpeta de cliente en red pública ──
                    res = ProcesarPV(res, rutaTemp, fileName);
                }
                //else if (prefijo == "26PC")
                //{
                //    // ── 26PC: buscar carpeta que empiece por el código en proveedorespedidos ──
                //    res = ProcesarPC(res, rutaTemp, fileName);
                //}
                //else if (prefijo == "26TR")
                //{
                //    // ── 26TR: buscar carpeta que empiece por el código en traspasos ──
                //    res = ProcesarTR(res, rutaTemp, fileName);
                //}
            }
            catch (Exception ex)
            {
                res.Error = ex.Message;
                MostrarError($"      Excepcion en {fileName}: {ex.Message}");
                EliminarArchivo(rutaTemp);
            }

            return res;
        }

        // ─────────────────────────────────────────────────────────────────────
        // 26PV — igual que antes: MySQL → cliente → carpeta en documentacion
        // ─────────────────────────────────────────────────────────────────────

        static ResultadoPdf ProcesarPV(ResultadoPdf res, string rutaTemp, string fileName)
        {
            res.NombreCliente = ObtenerNombreCliente(res.CodigoEncontrado!);

            if (res.NombreCliente == null)
            {
                EliminarArchivo(rutaTemp);
                res.Error = $"Código {res.CodigoEncontrado} no encontrado en BD (26PV)";
                return res;
            }

            // Subcarpeta: "FACTURA Fxxx" si existe, si no "PEDIDO 26PVxxx"
            string directoriofinal = "PEDIDO " + res.CodigoEncontrado;
            res.FacturaEncontrada  = ObtenerFacturaCliente(res.CodigoEncontrado!);
            if (!string.IsNullOrWhiteSpace(res.FacturaEncontrada))
                directoriofinal = "FACTURA " + res.FacturaEncontrada;

            string clienteLimpio  = LimpiarNombreCarpeta(res.NombreCliente.Replace(" ", ""));
            string carpetaDestino = Path.Combine(_cfg.RutaBaseRed, clienteLimpio, directoriofinal);

            CrearSiNoExiste(carpetaDestino);
            CopiarYEliminarTemp(rutaTemp, carpetaDestino, fileName, res);
            return res;
        }

        // ─────────────────────────────────────────────────────────────────────
        // 26PC — buscar carpeta que comience por el código en proveedorespedidos
        //        \\serveroversea\oversea2026\proveedorespedidos\
        // ─────────────────────────────────────────────────────────────────────

        static ResultadoPdf ProcesarPC(ResultadoPdf res, string rutaTemp, string fileName)
        {
            string raiz = Path.Combine(_cfg.ServidorOversea2026, "proveedorespedidos");

            string? carpetaDestino = BuscarCarpetaQueEmpieza(raiz, res.CodigoEncontrado!);

            if (carpetaDestino == null)
            {
                EliminarArchivo(rutaTemp);
                res.Error = $"[26PC] No existe carpeta para el codigo {res.CodigoEncontrado} en {raiz}";
                MostrarError($"      {res.Error}");
                return res;
            }

            Log($"      [26PC] Carpeta encontrada: {carpetaDestino}");
            CopiarYEliminarTemp(rutaTemp, carpetaDestino, fileName, res);
            return res;
        }

        // ─────────────────────────────────────────────────────────────────────
        // 26TR — buscar carpeta que comience por el código en traspasos
        //        \\serveroversea\oversea2026\traspasos\
        // ─────────────────────────────────────────────────────────────────────

        static ResultadoPdf ProcesarTR(ResultadoPdf res, string rutaTemp, string fileName)
        {
            string raiz = Path.Combine(_cfg.ServidorOversea2026, "traspasos");

            string? carpetaDestino = BuscarCarpetaQueEmpieza(raiz, res.CodigoEncontrado!);

            if (carpetaDestino == null)
            {
                EliminarArchivo(rutaTemp);
                res.Error = $"[26TR] No existe carpeta para el codigo {res.CodigoEncontrado} en {raiz}";
                MostrarError($"      {res.Error}");
                return res;
            }

            Log($"      [26TR] Carpeta encontrada: {carpetaDestino}");
            CopiarYEliminarTemp(rutaTemp, carpetaDestino, fileName, res);
            return res;
        }

        /// <summary>
        /// Busca dentro de <paramref name="raiz"/> la primera subcarpeta
        /// cuyo nombre empiece por <paramref name="codigo"/> (insensible a mayúsculas).
        /// Devuelve la ruta completa o null si no existe ninguna.
        /// </summary>
        static string? BuscarCarpetaQueEmpieza(string raiz, string codigo)
        {
            if (!Directory.Exists(raiz))
            {
                LogWarn($"      Ruta raiz no accesible: {raiz}");
                return null;
            }

            return Directory
                .EnumerateDirectories(raiz)
                .FirstOrDefault(d =>
                    Path.GetFileName(d).StartsWith(codigo, StringComparison.OrdinalIgnoreCase));
        }

        // ── UTILIDADES COMUNES ────────────────────────────────────────────────

        /// <summary>
        /// Copia el temporal al destino y lo elimina; actualiza el resultado.
        /// </summary>
        static void CopiarYEliminarTemp(string rutaTemp, string carpetaDestino, string fileName, ResultadoPdf res)
        {
            string rutaDestino = Path.Combine(carpetaDestino, fileName);
            // overwrite: true — si el archivo ya existe lo sobreescribe
            File.Copy(rutaTemp, rutaDestino, overwrite: true);
            EliminarArchivo(rutaTemp);

            res.RutaDestino = rutaDestino;
            res.Exitoso     = true;
        }

        static void CrearSiNoExiste(string carpeta)
        {
            if (!Directory.Exists(carpeta))
            {
                Directory.CreateDirectory(carpeta);
                Log($"      Carpeta creada: {carpeta}");
            }
        }

        static void EliminarArchivo(string ruta)
        {
            try { if (File.Exists(ruta)) File.Delete(ruta); }
            catch (Exception ex) { LogWarn($"No se pudo eliminar '{ruta}': {ex.Message}"); }
        }

        /// <summary>
        /// Elimina todos los PDFs dentro de una carpeta (deja la carpeta).
        /// </summary>
        static void LimpiarCarpeta(string carpeta)
        {
            if (!Directory.Exists(carpeta)) return;
            int eliminados = 0;
            foreach (string archivo in Directory.EnumerateFiles(carpeta, "*.pdf", SearchOption.TopDirectoryOnly))
            {
                EliminarArchivo(archivo);
                eliminados++;
            }
            // También limpiar archivos sin extensión o temporales con GUID
            foreach (string archivo in Directory.EnumerateFiles(carpeta, "*", SearchOption.TopDirectoryOnly))
            {
                EliminarArchivo(archivo);
                eliminados++;
            }
            if (eliminados > 0)
                Log($"Carpeta limpiada ({eliminados} archivos eliminados): {carpeta}");
        }

        static string? ExtraerCodigo(string texto)
        {
            var m = CodigoRegex.Match(texto);
            return m.Success ? m.Groups[1].Value.ToUpperInvariant() : null;
        }

        static string? BuscarCodigoEnPdf(string rutaPdf)
        {
            using var reader = new PdfReader(rutaPdf);
            using var doc    = new PdfDocument(reader);
            for (int i = 1; i <= doc.GetNumberOfPages(); i++)
            {
                string texto  = PdfTextExtractor.GetTextFromPage(doc.GetPage(i), new SimpleTextExtractionStrategy());
                string? codigo = ExtraerCodigo(texto);
                if (codigo != null) return codigo;
            }
            return null;
        }

        static string RutaUnica(string ruta)
        {
            if (!File.Exists(ruta)) return ruta;
            string dir   = Path.GetDirectoryName(ruta)!;
            string base_ = Path.GetFileNameWithoutExtension(ruta);
            string ext   = Path.GetExtension(ruta);
            int n = 1;
            string nueva;
            do { nueva = Path.Combine(dir, $"{base_}_{n++}{ext}"); }
            while (File.Exists(nueva));
            return nueva;
        }

        static string LimpiarNombreCarpeta(string nombre)
        {
            char[] invalidos = Path.GetInvalidFileNameChars();
            return string.Concat(nombre.Select(c => invalidos.Contains(c) ? '_' : c))
                         .Trim().TrimEnd('.');
        }

        // ── REGISTRO DE CORREOS PROCESADOS ──────────────────────────────────────
        // Guarda los EntryID de cada correo ya procesado en un archivo .txt
        // (un ID por línea). El archivo se rota por mes para no crecer indefinidamente.

        static void CargarRegistroProcesados()
        {
            string carpeta = !string.IsNullOrWhiteSpace(_cfg.CarpetaLog)
                ? _cfg.CarpetaLog
                : Path.Combine(AppContext.BaseDirectory, "logs");

            Directory.CreateDirectory(carpeta);

            // Un archivo por mes: procesados_yyyy-MM.txt
            string nombreArchivo = $"procesados_{DateTime.Now:yyyy-MM}.txt";
            _rutaRegistroProcesados = Path.Combine(carpeta, nombreArchivo);

            if (File.Exists(_rutaRegistroProcesados))
            {
                var ids = File.ReadAllLines(_rutaRegistroProcesados)
                              .Where(l => !string.IsNullOrWhiteSpace(l));
                foreach (var id in ids)
                    _correosYaProcesados.Add(id.Trim());

                Log($"Registro de procesados cargado: {_correosYaProcesados.Count} correos ({_rutaRegistroProcesados})");
            }
            else
            {
                Log($"Registro de procesados nuevo: {_rutaRegistroProcesados}");
            }
        }

        static void GuardarRegistroProcesados()
        {
            try
            {
                File.WriteAllLines(_rutaRegistroProcesados, _correosYaProcesados,
                                   System.Text.Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LogWarn($"No se pudo guardar el registro de procesados: {ex.Message}");
            }
        }

        // ── LOG ───────────────────────────────────────────────────────────────

        static void IniciarLog()
        {
            string carpeta = !string.IsNullOrWhiteSpace(_cfg.CarpetaLog)
                ? _cfg.CarpetaLog
                : Path.Combine(AppContext.BaseDirectory, "logs");

            Directory.CreateDirectory(carpeta);
            _rutaLog   = Path.Combine(carpeta, $"log_{DateTime.Now:yyyy-MM-dd}.txt");
            _logWriter = new StreamWriter(_rutaLog, append: true, System.Text.Encoding.UTF8)
                         { AutoFlush = true };
        }

        static void CerrarLog()
        {
            _logWriter?.Flush();
            _logWriter?.Close();
            _logWriter?.Dispose();
            _logWriter = null;
        }

        static void Escribir(string msg, string nivel = "INFO")
        {
            try { _logWriter?.WriteLine($"[{DateTime.Now:HH:mm:ss}] [{nivel,-5}] {msg}"); }
            catch { }
        }

        static void Log(string msg)      { if (_cfg.ModoDebug) Escribir(msg, "DEBUG"); }
        static void LogInfo(string msg)  => Escribir(msg, "INFO");
        static void LogOk(string msg)    => Escribir(msg, "OK");
        static void LogWarn(string msg)  => Escribir(msg, "WARN");
        static void MostrarError(string msg) => Escribir(msg, "ERROR");

        static void Titulo(string texto)
        {
            string sep = "".PadRight(65, '=');
            _logWriter?.WriteLine($"\n{sep}");
            _logWriter?.WriteLine($"  NUEVA SESION — {texto}");
            _logWriter?.WriteLine($"  Fecha : {DateTime.Now:dd/MM/yyyy HH:mm:ss}");
            _logWriter?.WriteLine($"  Log   : {_rutaLog}");
            _logWriter?.WriteLine($"{sep}\n");
        }

        static void MostrarResultado(ResultadoPdf r)
        {
            if (r.Exitoso)
            {
                LogOk($"    OK  {r.NombreArchivo}");
                LogOk($"        Codigo  : {r.CodigoEncontrado} (via {r.FuenteCodigo})");
                if (r.NombreCliente != null)
                    LogOk($"        Cliente : {r.NombreCliente}");
                LogOk($"        Destino : {r.RutaDestino}");
            }
            else if (r.CodigoEncontrado == null)
                LogWarn($"    ??  {r.NombreArchivo} — {r.Error}");
            else
            {
                MostrarError($"    XX  {r.NombreArchivo}");
                MostrarError($"        Codigo : {r.CodigoEncontrado} (via {r.FuenteCodigo})");
                MostrarError($"        {r.Error}");
            }
        }

        static void Resumen(int total, int ok, int sinCodigo, int sinCliente, int errores)
        {
            string sep = "".PadRight(50, '-');
            LogInfo(sep);
            LogInfo("  RESUMEN FINAL");
            LogInfo($"  PDFs procesados        : {total}");
            LogOk  ($"  Clasificados OK        : {ok}");
            LogWarn($"  Sin codigo encontrado  : {sinCodigo}");
            LogWarn($"  Codigo no en BD        : {sinCliente}");
            MostrarError($"  Otros errores          : {errores}");
            LogInfo(sep);
        }
    }
}
