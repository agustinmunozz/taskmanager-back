using Entity;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Reflection.Emit;

namespace BLL
{
    public static class Exportador
    {
        public static readonly string DirPath = "";
		//public static readonly string DirPath = new DAL.AppConfigurationHelper().Get().Path_Reportes;
		//public static readonly string DirIvr = new DAL.AppConfigurationHelper().Get().Path_Ivr;
		//public static readonly string DirPathCob = new DAL.AppConfigurationHelper().Get().Path_ReporteCobranza;
		//private static readonly DAL.Seguridad.UsuariosDA oUsuariosDA = new DAL.Seguridad.UsuariosDA();

		public enum Formato
        {
            MONEDA = 1,
            ENTERO = 2,
            AJUSTA_TEXTO = 3
        };

		public static DataTable ToDataTable<T>(List<T> items, Dictionary<string, string> properties)
		{
			DataTable dataTable = new DataTable(typeof(T).Name);
			PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
									   .Where(p => properties.ContainsKey(p.Name)).ToArray();
			foreach (PropertyInfo prop in Props)
			{
				//Setting column names as Property names  
				dataTable.Columns.Add(properties[prop.Name]);
			}
			foreach (T item in items)
			{
				var values = new object[Props.Length];
				for (int i = 0; i < Props.Length; i++)
				{

					values[i] = Props[i].GetValue(item, null);
				}
				dataTable.Rows.Add(values);
			}

			return dataTable;
		}


		#region Generador de Excel
		public static string GenerarExcel_NoUsar(DataTable Datos, string[] NombreColumnas, string[] DatosColumnas, string Titulo, string[] Filtros, string NombreArchivo)
        {
            try
            {
                //if (Datos == null || Datos.Rows.Count == 0)
                //    throw new EmptyObjectException();

                using (ExcelPackage Archivo = new ExcelPackage())
                {
                    ExcelWorksheet Hoja = Archivo.Workbook.Worksheets.Add(NombreArchivo.Length > 31 ? NombreArchivo.Substring(0, 28) + "..." : NombreArchivo);

                    #region Seteo Background White

                    Hoja.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Hoja.Cells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);

                    #endregion

                    #region Seteo de titulo y Cantidad de registros

                    int CantColumnasTitulo = NombreColumnas.Length <= 2 ? 2 : (NombreColumnas.Length - 2);
                    ExcelRange CeldaTitulo = Hoja.Cells[3, 1, 3, CantColumnasTitulo];
                    CeldaTitulo.Merge = true;
                    CeldaTitulo.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    CeldaTitulo.Value = Titulo;
                    CeldaTitulo.Style.Font.Bold = true;

                    ExcelRange CeldaCantidad = Hoja.Cells[3, CantColumnasTitulo + 1, 3, CantColumnasTitulo + 2];
                    CeldaCantidad.Merge = true;
                    CeldaCantidad.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    CeldaCantidad.Value = "Cantidad de Registros: " + Datos.Rows.Count;
                    CeldaCantidad.Style.Font.Bold = true;
                    CeldaCantidad.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    CeldaCantidad.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkOrange);
                    CeldaCantidad.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    CeldaCantidad.Style.Font.Size = 10;

                    #endregion

                    #region Filtros
                    int CantFiltros = 0;
                    if (Filtros != null && Filtros.Length > 0)
                    {
                        CantFiltros = Filtros.Length;

                        for (int i = 1; i <= CantFiltros; i++)
                        {
                            ExcelRange CeldaFiltro = Hoja.Cells[3 + i, 1, 3 + i, NombreColumnas.Length];
                            CeldaFiltro.Merge = true;
                            CeldaFiltro.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            CeldaFiltro.Value = Filtros[i - 1];
                        }
                    }

                    #endregion

                    if (Datos != null && Datos.Rows.Count != 0)
                    {
                        #region Header de grilla

                        for (int i = 0; i < NombreColumnas.Length; i++)
                        {
                            ExcelRange CeldaHeader = Hoja.Cells[5 + CantFiltros, i + 1];
                            CeldaHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CeldaHeader.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue);
                            CeldaHeader.Style.Font.Color.SetColor(System.Drawing.Color.White);
                            CeldaHeader.Style.Font.Size = 10;
                            CeldaHeader.Style.Font.Bold = true;
                            CeldaHeader.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            CeldaHeader.Value = NombreColumnas[i];
                        }

                        #endregion

                        #region DatosGrilla

                        int RowBase = 6 + CantFiltros;

                        if (DatosColumnas != null)
                        {
                            for (int i = 0; i < Datos.Rows.Count; i++)
                            {
                                for (int j = 0; j < DatosColumnas.Length; j++)
                                {


                                    ExcelRange CeldaGrilla = Hoja.Cells[RowBase + i, j + 1];

                                    CeldaGrilla.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    CeldaGrilla.Style.Font.Size = 10;
                                    CeldaGrilla.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    if (Datos.Columns[DatosColumnas[j]] != null)
                                    {
                                        if (Datos.Rows[i][DatosColumnas[j]].GetType() == typeof(DateTime))
                                            CeldaGrilla.Value = Convert.ToDateTime(Datos.Rows[i][DatosColumnas[j]]).ToString();
                                        else
                                            CeldaGrilla.Value = Datos.Rows[i][DatosColumnas[j]];
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < Datos.Rows.Count; i++)
                            {
                                for (int j = 0; j < Datos.Columns.Count; j++)
                                {
                                    ExcelRange CeldaGrilla = Hoja.Cells[RowBase + i, j + 1];

                                    CeldaGrilla.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    CeldaGrilla.Style.Font.Size = 10;
                                    CeldaGrilla.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                    if (Datos.Rows[i][j].GetType() == typeof(DateTime))
                                        CeldaGrilla.Value = Convert.ToDateTime(Datos.Rows[i][j]).ToString();
                                    else
                                        CeldaGrilla.Value = Datos.Rows[i][j];
                                }
                            }
                        }

                        #endregion
                    }
                    #region Guardar Archivo
                    Hoja.Cells.AutoFitColumns();
                    Hoja.Protection.IsProtected = false;
                    Hoja.Protection.AllowSelectLockedCells = false;
                    string PathArchivo = NombreArchivo + ".xlsx";
                    Archivo.SaveAs(new FileInfo(DirPath + PathArchivo));
                    return PathArchivo;

                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GenerarExcel(DataTable dtSrc, string[] NombreColumnas, string Titulo, string NombreArchivo, string path = null, string[] orden = null, long? UsuarioLogin = null, Dictionary<int, Formato> Formato = null)
        {
            try
            {
                //if (dtSrc == null || dtSrc.Rows.Count == 0)
                //    throw new EmptyObjectException();
                if (orden != null)
                {
                    OrderColumm(dtSrc, orden);
                }
                using (ExcelPackage Archivo = new ExcelPackage())
                {
                    ExcelWorksheet Hoja = Archivo.Workbook.Worksheets.Add(NombreArchivo.Length > 31 ? NombreArchivo.Substring(0, 28) + "..." : NombreArchivo);
                    #region Seteo Background White

                    IEnumerable<int> dateColumns = from DataColumn d in dtSrc.Columns
                                                   where d.DataType == typeof(DateTime) ||
                                                   d.ColumnName.Contains("Date")
                                                   select d.Ordinal + 1;
                    foreach (int dc in dateColumns)
                    {
                        Hoja.Cells[5, dc, dtSrc.Rows.Count + 6, dc].Style.Numberformat.Format = "dd/MM/yyyy";
                    }

                    Hoja.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Hoja.Cells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                    #endregion

                    int CantColumnasTitulo = NombreColumnas.Length <= 2 ? 2 : (NombreColumnas.Length - 2);

                    #region Seteo Fecha y usuario
                    if (UsuarioLogin != null)
                    {
                        ExcelRange CeldaFecha = Hoja.Cells[1, CantColumnasTitulo + 1, 1, CantColumnasTitulo + 2];
                        CeldaFecha.Merge = true;
                        CeldaFecha.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        CeldaFecha.Value = "Fecha: " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
                        CeldaFecha.Style.Font.Bold = true;
                        CeldaFecha.Style.Font.Size = 10;

                        //Entity.Models.User oUer = .Get<Entity.Models.User>(UsuarioLogin.Value);

                        ExcelRange CeldaUsuario = Hoja.Cells[2, CantColumnasTitulo + 1, 2, CantColumnasTitulo + 2];
                        CeldaUsuario.Merge = true;
                        CeldaUsuario.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        //CeldaUsuario.Value = "Usuario: " + (oUsuarios != null ? oUsuarios.Usuario : "");
                        CeldaUsuario.Style.Font.Bold = true;
                        CeldaUsuario.Style.Font.Size = 10;
                    }
                    #endregion Seteo Fecha y usuario

                    #region Seteo de titulo y Cantidad de registros

                    ExcelRange CeldaTitulo = Hoja.Cells[3, 1, 3, CantColumnasTitulo];
                    CeldaTitulo.Merge = true;
                    CeldaTitulo.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    CeldaTitulo.Value = Titulo;
                    CeldaTitulo.Style.Font.Bold = true;

                    ExcelRange CeldaCantidad = Hoja.Cells[3, CantColumnasTitulo + 1, 3, CantColumnasTitulo + 2];
                    CeldaCantidad.Merge = true;
                    CeldaCantidad.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    CeldaCantidad.Value = "Cantidad de Registros: " + dtSrc.Rows.Count;
                    CeldaCantidad.Style.Font.Bold = true;
                    CeldaCantidad.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkOrange);
                    CeldaCantidad.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    CeldaCantidad.Style.Font.Size = 10;
                    #endregion

                    #region Header de grilla
                    for (int i = 0; i < NombreColumnas.Length; i++)
                    {
                        ExcelRange CeldaHeader = Hoja.Cells[5, i + 1];
                        CeldaHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CeldaHeader.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue);
                        CeldaHeader.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        CeldaHeader.Style.Font.Size = 10;
                        CeldaHeader.Style.Font.Bold = true;
                        CeldaHeader.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        CeldaHeader.Value = NombreColumnas[i];
                    }
                    #endregion
                    Hoja.Cells["A6"].LoadFromDataTable(dtSrc, ((NombreColumnas.Length > 0) ? false : true));

                    (from DataColumn d in dtSrc.Columns select d.Ordinal + 1).ToList().ForEach(dc =>
                    {
                        //borde
                        Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        //alineacion
                        Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //funete
                        Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].Style.Font.Size = 10;
                        //Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].AutoFitColumns();

                    });

                    #region Formato de columnas

                    if (Formato != null && Formato.Count > 0)
                    {
                        foreach (var item in Formato)
                        {
                            switch (item.Value)
                            {
                                case Exportador.Formato.MONEDA:
                                    Hoja.Cells[6, item.Key, dtSrc.Rows.Count + 5, item.Key].Style.Numberformat.Format = "$#,##0.00";
                                    Hoja.Cells[6, item.Key, dtSrc.Rows.Count + 5, item.Key].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;                                    
                                    break;
                                default:
                                    break;
                            }
                        }
                    }

                    #endregion

                    #region Guardar Archivo
                    Hoja.Cells.AutoFitColumns();
                    Hoja.Protection.IsProtected = false;
                    Hoja.Protection.AllowSelectLockedCells = false;
                    string PathArchivo = NombreArchivo + ".xlsx";
                    Archivo.SaveAs(new FileInfo(path != null ? path + PathArchivo : DirPath + PathArchivo));
                    Hoja.Dispose();
                    Archivo.Dispose();

                    return PathArchivo;

                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        public static string GenerarExcel2(DataTable dtSrc, string[] NombreColumnas, string[] Titulos, string NombreArchivo, string path = null, string[] orden = null, long? UsuarioLogin = null, Dictionary<int, Formato> Formato = null, bool IsProtected = false)
        {
           try
            {
                //if (dtSrc == null || dtSrc.Rows.Count == 0)
                //    throw new EmptyObjectException();
                if (orden != null)
                {
                    OrderColumm(dtSrc, orden);
                }
                using (ExcelPackage Archivo = new ExcelPackage())
                {
                    ExcelWorksheet Hoja = Archivo.Workbook.Worksheets.Add(NombreArchivo.Length > 31 ? NombreArchivo.Substring(0, 28) + "..." : NombreArchivo);
                    #region Seteo Background White

                    IEnumerable<int> dateColumns = from DataColumn d in dtSrc.Columns
                                                   where d.DataType == typeof(DateTime) ||
                                                   d.ColumnName.Contains("Date")
                                                   select d.Ordinal + 1;
                    foreach (int dc in dateColumns)
                    {
                        Hoja.Cells[5, dc, dtSrc.Rows.Count + 6 + Titulos.Length, dc].Style.Numberformat.Format = "dd/MM/yyyy";
                    }

                    Hoja.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Hoja.Cells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                    #endregion

                    int CantColumnasTitulo = NombreColumnas.Length <= 2 ? 2 : (NombreColumnas.Length - 2);

                    #region Seteo Fecha y usuario
                    if (UsuarioLogin != null)
                    {
                        ExcelRange CeldaFecha = Hoja.Cells[1, CantColumnasTitulo + 1, 1, CantColumnasTitulo + 2];
                        CeldaFecha.Merge = true;
                        CeldaFecha.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        CeldaFecha.Value = "Fecha: " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss");
                        CeldaFecha.Style.Font.Bold = true;
                        CeldaFecha.Style.Font.Size = 10;

                        //Entity.Seguridad.Usuarios oUsuarios = oUsuariosDA.Get<Entity.Seguridad.Usuarios>(UsuarioLogin.Value);

                        ExcelRange CeldaUsuario = Hoja.Cells[2, CantColumnasTitulo + 1, 2, CantColumnasTitulo + 2];
                        CeldaUsuario.Merge = true;
                        CeldaUsuario.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        //CeldaUsuario.Value = "Usuario: " + (oUsuarios != null ? oUsuarios.Usuario : "");
                        CeldaUsuario.Style.Font.Bold = true;
                        CeldaUsuario.Style.Font.Size = 10;
                    }
                    #endregion Seteo Fecha y usuario

                    #region Seteo de titulo y Cantidad de registros
                    int CantTitulo = Titulos.Length <= 2 ? 2 : (Titulos.Length - 2);
                    for (int i = 0; i < Titulos.Length; i++)
                    {
                        ExcelRange CeldaTitulo = Hoja.Cells[3 + i, 1, 3 + i, CantColumnasTitulo];
                        CeldaTitulo.Merge = true;
                        CeldaTitulo.Style.HorizontalAlignment = (i == 0 ? ExcelHorizontalAlignment.Center : ExcelHorizontalAlignment.Left);
                        CeldaTitulo.Value = Titulos[i];
                        CeldaTitulo.Style.Font.Bold = true;
                    }

                    ExcelRange CeldaCantidad = Hoja.Cells[3, CantColumnasTitulo + 1, 3, CantColumnasTitulo + 2];
                    CeldaCantidad.Merge = true;
                    CeldaCantidad.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    CeldaCantidad.Value = "Cantidad de Registros: " + dtSrc.Rows.Count;
                    CeldaCantidad.Style.Font.Bold = true;
                    CeldaCantidad.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkOrange);
                    CeldaCantidad.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    CeldaCantidad.Style.Font.Size = 10;
                    #endregion

                    #region Header de grilla
                    for (int i = 0; i < NombreColumnas.Length; i++)
                    {
                        ExcelRange CeldaHeader = Hoja.Cells[5 + CantTitulo + 1, i + 1];
                        CeldaHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CeldaHeader.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue);
                        CeldaHeader.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        CeldaHeader.Style.Font.Size = 10;
                        CeldaHeader.Style.Font.Bold = true;
                        CeldaHeader.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        CeldaHeader.Value = NombreColumnas[i];
                    }
                    #endregion
                    Hoja.Cells["A" + (6 + CantTitulo + 1).ToString()].LoadFromDataTable(dtSrc, ((NombreColumnas.Length > 0) ? false : true));

                    (from DataColumn d in dtSrc.Columns select d.Ordinal + 1).ToList().ForEach(dc =>
                    {
                        //borde
                        Hoja.Cells[6 + CantTitulo, dc, dtSrc.Rows.Count + 6 + CantTitulo, dc].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Hoja.Cells[6 + CantTitulo, dc, dtSrc.Rows.Count + 6 + CantTitulo, dc].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        //alineacion
                        Hoja.Cells[6 + CantTitulo, dc, dtSrc.Rows.Count + 6 + CantTitulo, dc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //funete
                        Hoja.Cells[6 + CantTitulo, dc, dtSrc.Rows.Count + 6 + CantTitulo, dc].Style.Font.Size = 10;
                    });

                    #region Formato de columnas

                    if (Formato != null && Formato.Count > 0)
                    {
                        foreach (var item in Formato)
                        {
                            switch (item.Value)
                            {
                                case Exportador.Formato.MONEDA:
                                    Hoja.Cells[6 + CantTitulo, item.Key, dtSrc.Rows.Count + 6 + CantTitulo, item.Key].Style.Numberformat.Format = "$#,##0.00";
                                    Hoja.Cells[6 + CantTitulo, item.Key, dtSrc.Rows.Count + 6 + CantTitulo, item.Key].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                    break;
                                case Exportador.Formato.AJUSTA_TEXTO:
                                    Hoja.Cells[6 + CantTitulo, item.Key, dtSrc.Rows.Count + 6 + CantTitulo, item.Key].Style.WrapText = true;
                                    break;
                                default:
                                    break;
                            }
                        }
                    }

                    #endregion

                    #region Guardar Archivo
                    Hoja.Cells.AutoFitColumns();
                    Hoja.Protection.IsProtected = IsProtected;
                    Hoja.Protection.AllowSelectLockedCells = false;
                    string PathArchivo = NombreArchivo + ".xlsx";
                    Archivo.SaveAs(new FileInfo(path != null ? path + PathArchivo : DirPath + PathArchivo));
                    return PathArchivo;

                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GenerarExcel(DataTable dtSrc, string[] NombreColumnas, string[] Titulos, string NombreArchivo, string FormatoFecha = "dd/MM/yyyy")
        {
            try
            {
                //if (dtSrc == null || dtSrc.Rows.Count == 0)
                //    throw new EmptyObjectException();

                using (ExcelPackage Archivo = new ExcelPackage())
                {
                    ExcelWorksheet Hoja = Archivo.Workbook.Worksheets.Add(NombreArchivo.Length > 31 ? NombreArchivo.Substring(0, 28) + "..." : NombreArchivo);
                    #region Seteo Background White

                    IEnumerable<int> dateColumns = from DataColumn d in dtSrc.Columns
                                                   where d.DataType == typeof(DateTime) ||
                                                   d.ColumnName.Contains("Date")
                                                   select d.Ordinal + 1;
                    foreach (int dc in dateColumns)
                    {
                        Hoja.Cells[4 + Titulos.Length, dc, dtSrc.Rows.Count + 5 + Titulos.Length, dc].Style.Numberformat.Format = FormatoFecha;
                    }

                    Hoja.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Hoja.Cells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                    #endregion

                    #region Seteo de titulo y Cantidad de registros

                    int CantColumnasTitulo = NombreColumnas.Length <= 2 ? 2 : (NombreColumnas.Length - 2);
                    for (int i = 0; i < Titulos.Length; i++)
                    {
                        ExcelRange CeldaTitulo = Hoja.Cells[3 + i, 1, 3 + i, CantColumnasTitulo];
                        CeldaTitulo.Merge = true;
                        CeldaTitulo.Style.HorizontalAlignment = (i == 0 ? ExcelHorizontalAlignment.Center : ExcelHorizontalAlignment.Left);
                        CeldaTitulo.Value = Titulos[i];
                        CeldaTitulo.Style.Font.Bold = true;
                    }

                    ExcelRange CeldaCantidad = Hoja.Cells[3, CantColumnasTitulo + 1, 3, CantColumnasTitulo + 2];
                    CeldaCantidad.Merge = true;
                    CeldaCantidad.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    CeldaCantidad.Value = "Cantidad de Registros: " + dtSrc.Rows.Count;
                    CeldaCantidad.Style.Font.Bold = true;
                    CeldaCantidad.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkOrange);
                    CeldaCantidad.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    CeldaCantidad.Style.Font.Size = 10;
                    #endregion

                    #region Header de grilla
                    for (int i = 0; i < NombreColumnas.Length; i++)
                    {
                        ExcelRange CeldaHeader = Hoja.Cells[4 + Titulos.Length, i + 1];
                        CeldaHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        CeldaHeader.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue);
                        CeldaHeader.Style.Font.Color.SetColor(System.Drawing.Color.White);
                        CeldaHeader.Style.Font.Size = 10;
                        CeldaHeader.Style.Font.Bold = true;
                        CeldaHeader.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        CeldaHeader.Value = NombreColumnas[i];
                    }
                    #endregion
                    int columnDatos = 5 + Titulos.Length;
                    Hoja.Cells["A" + columnDatos.ToString()].LoadFromDataTable(dtSrc, ((NombreColumnas.Length > 0) ? false : true));

                    (from DataColumn d in dtSrc.Columns select d.Ordinal + 1).ToList().ForEach(dc =>
                    {
                        //borde
                        Hoja.Cells[5 + Titulos.Length, dc, dtSrc.Rows.Count + 4 + Titulos.Length, dc].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        Hoja.Cells[5 + Titulos.Length, dc, dtSrc.Rows.Count + 4 + Titulos.Length, dc].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        //alineacion
                        Hoja.Cells[5 + Titulos.Length, dc, dtSrc.Rows.Count + 4 + Titulos.Length, dc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //funete
                        Hoja.Cells[5 + Titulos.Length, dc, dtSrc.Rows.Count + 4 + Titulos.Length, dc].Style.Font.Size = 10;
                        //Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].AutoFitColumns();
                    });
                    #region Guardar Archivo
                    Hoja.Cells.AutoFitColumns(15, 150);
                    Hoja.Protection.IsProtected = false;
                    Hoja.Protection.AllowSelectLockedCells = false;
                    string PathArchivo = NombreArchivo + ".xlsx";
                    Archivo.SaveAs(new FileInfo(DirPath + PathArchivo));
                    return PathArchivo;

                    #endregion
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public static string GenerarExcel(DataSet dsSrc, string[][] TitulosColumnas, string[][] Titulos, string NombreArchivo, List<Dictionary<int, Formato>> LstFormatoCol = null)
        {
            try
            {
                using (ExcelPackage Archivo = new ExcelPackage())
                {
                    ExcelWorksheet Hoja = Archivo.Workbook.Worksheets.Add(NombreArchivo.Length > 31 ? NombreArchivo.Substring(0, 28) + "..." : NombreArchivo);
                    int pos = 0;//suma cantidad de datatable
                    bool priVacia = false;
                    Hoja.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Hoja.Cells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                    // variables para determinar la posicion y rango - si se quiere mejorar usar el piensa piensa..!
                    int posRango = 0; int posRango4 = 4; int posRango3 = 3; int tablaAnterior = 0; int tablaActual = 0;                    
                    foreach (DataTable dtSrc in dsSrc.Tables)
                    {                        
                        //reasignamos la cantidad de registros del dataset anterior
                        tablaAnterior = tablaActual;
                        tablaActual = dtSrc.Rows.Count;
                        if (dtSrc.Rows.Count > 0)
                        {
                            #region Seteo Background White

                            int columnDatos = 5 + Titulos[pos].Length;
                            if (pos == 0 || priVacia == true)
                            {

                                posRango = columnDatos + ((priVacia) ? 1 : 0);
                                posRango4 = 4 + Titulos[pos].Length;
                                priVacia = false;
                            }
                            else
                            {
                                //posRango4 = posRango4 + columnDatos + dtSrc.Rows.Count;
                                //posRango3 = posRango3 + columnDatos + dtSrc.Rows.Count;
                                //posRango += columnDatos + dtSrc.Rows.Count + 1;

                                posRango4 = posRango4 + columnDatos + tablaAnterior;
                                posRango3 = posRango3 + columnDatos + tablaAnterior;
                                posRango += columnDatos + tablaAnterior + 1;
                            }

                            IEnumerable<int> dateColumns = from DataColumn d in dtSrc.Columns
                                                           where d.DataType == typeof(DateTime) ||
                                                           d.ColumnName.Contains("Date")
                                                           select d.Ordinal + 1;
                            foreach (int dc in dateColumns)
                            {
                                Hoja.Cells[posRango4, dc, dtSrc.Rows.Count + posRango, dc].Style.Numberformat.Format = "dd/MM/yyyy";
                            }


                            #endregion

                            #region Seteo de titulo y Cantidad de registros

                            int CantColumnasTitulo = TitulosColumnas[pos].Length <= 2 ? 2 : (TitulosColumnas[pos].Length - 2);

                            for (int i = 0; i < Titulos[pos].Length; i++)
                            {
                                ExcelRange CeldaTitulo = Hoja.Cells[posRango3 + pos + i, 1, posRango3 + pos + i, CantColumnasTitulo];
                                CeldaTitulo.Merge = true;
                                CeldaTitulo.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                CeldaTitulo.Value = Titulos[pos][i];
                                CeldaTitulo.Style.Font.Bold = true;
                            }

                            ExcelRange CeldaCantidad = Hoja.Cells[posRango3 + pos, CantColumnasTitulo + 1, posRango3 + pos, CantColumnasTitulo + 2];
                            CeldaCantidad.Merge = true;
                            CeldaCantidad.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            CeldaCantidad.Value = "Cantidad de Registros: " + dtSrc.Rows.Count;
                            CeldaCantidad.Style.Font.Bold = true;
                            CeldaCantidad.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkOrange);
                            CeldaCantidad.Style.Font.Color.SetColor(System.Drawing.Color.White);
                            CeldaCantidad.Style.Font.Size = 10;
                            #endregion

                            #region Header de grilla
                            for (int vHe = 0; vHe < TitulosColumnas[pos].Length; vHe++)
                            {
                                ExcelRange CeldaHeader = Hoja.Cells[posRango4 + pos, vHe + 1];
                                CeldaHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                CeldaHeader.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue);
                                CeldaHeader.Style.Font.Color.SetColor(System.Drawing.Color.White);
                                CeldaHeader.Style.Font.Size = 10;
                                CeldaHeader.Style.Font.Bold = true;
                                CeldaHeader.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                CeldaHeader.Value = TitulosColumnas[pos][vHe];
                            }
                            #endregion

                            Hoja.Cells["A" + posRango.ToString()].LoadFromDataTable(dtSrc, ((TitulosColumnas[pos].Length > 0) ? false : true));

                            (from DataColumn d in dtSrc.Columns select d.Ordinal + 1).ToList().ForEach(dc =>
                            {
                                //borde
                                Hoja.Cells[posRango, dc, dtSrc.Rows.Count + posRango - 1, dc].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                                Hoja.Cells[posRango, dc, dtSrc.Rows.Count + posRango - 1, dc].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                                //alineacion
                                Hoja.Cells[posRango, dc, dtSrc.Rows.Count + posRango, dc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                //funete
                                Hoja.Cells[posRango, dc, dtSrc.Rows.Count + posRango, dc].Style.Font.Size = 10;
                            });

                            if (LstFormatoCol != null && LstFormatoCol.Count > 0)
                            {
                                foreach (var item in LstFormatoCol[pos])
                                {
                                    switch (item.Value)
                                    {
                                        case Exportador.Formato.MONEDA:
                                            Hoja.Cells[posRango, item.Key, dtSrc.Rows.Count + posRango, item.Key].Style.Numberformat.Format = "$#,##0.00";
                                            Hoja.Cells[posRango, item.Key, dtSrc.Rows.Count + posRango, item.Key].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                            break;
                                        case Exportador.Formato.ENTERO:
                                            Hoja.Cells[posRango, item.Key, dtSrc.Rows.Count + posRango, item.Key].Style.Numberformat.Format = "#,##0";
                                            Hoja.Cells[posRango, item.Key, dtSrc.Rows.Count + posRango, item.Key].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                                            break;
                                        default:
                                            break;
                                    }
                                }
                            }                            
                        }
                        else
                        {
                            if (pos == 0)
                                priVacia = true;
                        }
                        pos++;                        
                    }

                    #region Guardar Archivo
                    Hoja.Cells.AutoFitColumns();
                    Hoja.Protection.IsProtected = false;
                    Hoja.Protection.AllowSelectLockedCells = false;
                    string PathArchivo = NombreArchivo + ".xlsx";
                    Archivo.SaveAs(new FileInfo(DirPath + PathArchivo));
                    return PathArchivo;

                    #endregion
                }
            }
            catch
            {
                return string.Empty;
            }
        }
        //public static string GenerateExcelCobranza(DataTable dtSrc, string[] NombreColumnas, string Titulo, string NombreArchivo)
        //{
        //    try
        //    {
        //        //if (dtSrc == null || dtSrc.Rows.Count == 0)
        //        //    throw new EmptyObjectException();

        //        using (ExcelPackage Archivo = new ExcelPackage())
        //        {
        //            ExcelWorksheet Hoja = Archivo.Workbook.Worksheets.Add(NombreArchivo.Length > 31 ? NombreArchivo.Substring(0, 28) + "..." : NombreArchivo);
        //            #region Seteo Background White

        //            IEnumerable<int> dateColumns = from DataColumn d in dtSrc.Columns
        //                                           where d.DataType == typeof(DateTime) ||
        //                                           d.ColumnName.Contains("Date")
        //                                           select d.Ordinal + 1;
        //            foreach (int dc in dateColumns)
        //            {
        //                Hoja.Cells[5, dc, dtSrc.Rows.Count + 6, dc].Style.Numberformat.Format = "dd/MM/yyyy";
        //            }

        //            Hoja.Cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //            Hoja.Cells.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
        //            #endregion

        //            #region Seteo de titulo y Cantidad de registros
        //            int CantColumnasTitulo = 0;
        //            if (NombreColumnas.Length > 0)
        //                CantColumnasTitulo = NombreColumnas.Length <= 2 ? 2 : (NombreColumnas.Length - 2);
        //            else
        //                CantColumnasTitulo = dtSrc.Columns.Count <= 2 ? 2 : (dtSrc.Columns.Count - 2);
        //            ExcelRange CeldaTitulo = Hoja.Cells[3, 1, 3, CantColumnasTitulo];
        //            CeldaTitulo.Merge = true;
        //            CeldaTitulo.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            CeldaTitulo.Value = Titulo;
        //            CeldaTitulo.Style.Font.Bold = true;

        //            ExcelRange CeldaCantidad = Hoja.Cells[3, CantColumnasTitulo + 1, 3, CantColumnasTitulo + 2];
        //            CeldaCantidad.Merge = true;
        //            CeldaCantidad.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //            CeldaCantidad.Value = "Cantidad de Registros: " + dtSrc.Rows.Count;
        //            CeldaCantidad.Style.Font.Bold = true;
        //            CeldaCantidad.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.DarkOrange);
        //            CeldaCantidad.Style.Font.Color.SetColor(System.Drawing.Color.White);
        //            CeldaCantidad.Style.Font.Size = 10;
        //            #endregion

        //            #region Header de grilla
        //            if (NombreColumnas.Length > 0)
        //            {
        //                for (int i = 0; i < NombreColumnas.Length; i++)
        //                {
        //                    ExcelRange CeldaHeader = Hoja.Cells[5, i + 1];
        //                    CeldaHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //                    CeldaHeader.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue);
        //                    CeldaHeader.Style.Font.Color.SetColor(System.Drawing.Color.White);
        //                    CeldaHeader.Style.Font.Size = 10;
        //                    CeldaHeader.Style.Font.Bold = true;
        //                    CeldaHeader.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        //                    CeldaHeader.Value = NombreColumnas[i];
        //                }
        //            }
        //            else
        //            {
        //                for (int i = 0; i < dtSrc.Columns.Count; i++)
        //                {
        //                    ExcelRange CeldaHeader = Hoja.Cells[5, i + 1];
        //                    CeldaHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //                    CeldaHeader.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Blue);
        //                    CeldaHeader.Style.Font.Color.SetColor(System.Drawing.Color.White);
        //                    CeldaHeader.Style.Font.Size = 10;
        //                    CeldaHeader.Style.Font.Bold = true;
        //                    CeldaHeader.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        //                    CeldaHeader.Value = dtSrc.Columns[i].ColumnName;
        //                }
        //            }
        //            #endregion
        //            Hoja.Cells["A6"].LoadFromDataTable(dtSrc, false);

        //            (from DataColumn d in dtSrc.Columns select d.Ordinal + 1).ToList().ForEach(dc =>
        //            {
        //                //borde
        //                Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        //                Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        //                //alineacion
        //                Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //                //funete
        //                Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].Style.Font.Size = 10;
        //                //Hoja.Cells[6, dc, dtSrc.Rows.Count + 5, dc].AutoFitColumns();
        //            });
        //            #region Guardar Archivo
        //            Hoja.Cells.AutoFitColumns();
        //            Hoja.Protection.IsProtected = false;
        //            Hoja.Protection.AllowSelectLockedCells = false;
        //            string PathArchivo = NombreArchivo + ".xlsx";
        //            //Archivo.SaveAs(new FileInfo(DirPathCob + PathArchivo));
        //            //return DirPathCob + PathArchivo;

        //            #endregion
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }

        //}
       
        //public static string GenerarExcelCSV(DataTable dtSrc, string NombreArchivo, bool isTXT = false)
        //{
        //    try
        //    {
        //        if (dtSrc == null || dtSrc.Rows.Count == 0)
        //            //throw new EmptyObjectException();

        //        StringBuilder sb = new StringBuilder();

        //        IEnumerable<string> columnNames = dtSrc.Columns.Cast<DataColumn>().
        //                                          Select(column => column.ColumnName);
        //        sb.AppendLine(string.Join(";", columnNames));

        //        foreach (DataRow row in dtSrc.Rows)
        //        {
        //            IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString().Replace("\n", ""));
        //            sb.AppendLine(string.Join(";", fields));
        //        }

        //        string pathArchivo = isTXT ? NombreArchivo + ".txt" : NombreArchivo + ".csv";

        //        var file = new FileInfo(DirPath + pathArchivo);

        //        StreamWriter sw = new StreamWriter(file.CreateText().BaseStream, Encoding.UTF8);
        //        sw.Write(sb.ToString());
        //        sw.Close();

        //        return pathArchivo;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //}
        //#endregion

        #region Exportadores
        public static DataTable ExcelToDataTable(string path, bool ContieneEncabezado)
        {
            var dt = new DataTable();
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = System.IO.File.OpenRead(DirPath + path))
                {
                    pck.Load(stream);
                }

                ExcelWorksheet worksheet = pck.Workbook.Worksheets[1];

                for (int PosCol = worksheet.Dimension.Start.Column; PosCol <= worksheet.Dimension.End.Column; PosCol++)
                {
                    string columnName = "Column " + PosCol;
                    var excelCell = worksheet.Cells[1, PosCol].Value;

                    if (excelCell != null)
                    {
                        var excelCellDataType = excelCell;
                        if (ContieneEncabezado == true)
                        {
                            excelCellDataType = worksheet.Cells[2, PosCol].Value;

                            columnName = excelCell.ToString().Trim();
                            if (dt.Columns.Contains(columnName) == true)
                            {
                                columnName = columnName + "_" + PosCol;
                            }
                        }
                        if (excelCellDataType is DateTime)
                        {
                            dt.Columns.Add(columnName, typeof(DateTime));
                        }
                        else if (excelCellDataType is Boolean)
                        {
                            dt.Columns.Add(columnName, typeof(Boolean));
                        }
                        else if (excelCellDataType is Double)
                        {
                            dt.Columns.Add(columnName, typeof(Decimal));
                        }
                        else
                        {
                            dt.Columns.Add(columnName, typeof(String));
                        }
                    }
                    else
                    {
                        dt.Columns.Add(columnName, typeof(String));
                    }
                }
                for (int vRow = worksheet.Dimension.Start.Row + Convert.ToInt32(ContieneEncabezado); vRow <= worksheet.Dimension.End.Row; vRow++)
                {
                    DataRow row = dt.NewRow();
                    for (int vCol = worksheet.Dimension.Start.Column; vCol <= worksheet.Dimension.End.Column; vCol++)
                    {
                        var excelCell = worksheet.Cells[vRow, vCol].Value;
                        if (excelCell != null)
                        {
                            try
                            {
                                row[vCol - 1] = excelCell;
                            }
                            catch
                            {
                            }
                        }
                    }
                    dt.Rows.Add(row);
                }
            }

            return dt;

        }
        public static DataTable TxtoDataTable(string path, char? separadorArchivos, bool firstLine)
        {
            try
            {
                DataTable dt = null;
                string[] encabezados;
                string[] lineaActual;
                using (StreamReader sr = File.OpenText(DirPath + path))
                {
                    string linea;

                    ArrayList lineas = new ArrayList();
                    //Lee el archivo de texto linea a linea
                    while (!sr.EndOfStream)
                    {
                        linea = sr.ReadLine();
                        lineas.Add(linea);
                    }

                    dt = new DataTable();

                    if (firstLine)
                    {
                        encabezados = lineas[0].ToString().Split(Convert.ToChar(separadorArchivos));
                        foreach (string col in encabezados)
                        {
                            dt.Columns.Add(col);
                        }
                    }

                    for (int p = 1; p < lineas.Count; p++)
                    {
                        lineaActual = lineas[p].ToString().Split(Convert.ToChar(separadorArchivos));
                        if (lineaActual.Length == dt.Columns.Count)
                        {
                            DataRow row = dt.NewRow();
                            row.ItemArray = lineaActual;
                            dt.Rows.Add(row);
                        }

                        else
                        {
                            throw new Exception();
                        }
                    }

                    sr.Close();

                }
                File.Delete(DirPath + path);
                return dt;
            }
            catch
            {

                throw new Exception("El número de datos no corresponde al numero de columnas");
            }

        }
        public static void OrderColumm(this DataTable dTable, String[] columnNames)
        {
            for (Int32 row = dTable.Columns.Count - 1; row >= 0; row--)
            {
                Boolean existeCol = false;
                foreach (var columnName in columnNames)
                {
                    if (dTable.Columns[row].ToString() == columnName)
                    {
                        existeCol = true;
                        break;
                    }
                }
                if (existeCol == false)
                    dTable.Columns.Remove(dTable.Columns[row]);
            }

            int columnIndex = 0;
            foreach (var columnName in columnNames)
            {
                dTable.Columns[columnName].SetOrdinal(columnIndex);
                columnIndex++;
            }
        }
        #endregion      
    }
}
#endregion