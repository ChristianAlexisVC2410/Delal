using Models;
using OfficeOpenXml;
using System.IO;

namespace BussinessLogic
{
    public class DelalLogic
    {
        public async Task<(object, MemoryStream)> CargarArchivoLogic(dynamic files)
        {
            try
            {
                if (files == null || files.Length == 0)
                {
                    throw new Exception("Solicitud incorrecta, no agregó el archivo");
                }

                List<Autobus> listBusOld= new List<Autobus> ();
                List<Autobus> listBusNew= new List<Autobus> ();
                List<Autobus> listBusRepetidos= new List<Autobus> ();
                List<Autobus> listBusAgregados= new List<Autobus> ();
                List<Autobus> listBusEliminados= new List<Autobus> ();
                List<Autobus> listBusFinalAux= new List<Autobus> ();
                List<Autobus> listBusFinal= new List<Autobus> ();
                MemoryStream memoryaGlobal= new MemoryStream ();
                object dataExcel = null;

                using (var stream = new MemoryStream())
                {
                    await files.CopyToAsync(stream);
                    stream.Position = 0;
                    using (var package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        int col = 2;
                        Autobus bus = null;


                        for (int row = 3; worksheet.Cells[row, col].Value != null; row++)
                        {
                            if (//(worksheet.Cells[row, 1].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 1].Value.ToString())) ||
                            (worksheet.Cells[row, 2].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 2].Value.ToString())) ||
                            (worksheet.Cells[row, 3].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 3].Value.ToString())) ||
                            (worksheet.Cells[row, 4].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 4].Value.ToString())) ||
                            (worksheet.Cells[row, 5].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 5].Value.ToString())) ||
                            (worksheet.Cells[row, 6].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 6].Value.ToString())) ||
                            (worksheet.Cells[row, 7].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 7].Value.ToString())) ||
                            (worksheet.Cells[row, 8].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 8].Value.ToString())) ||
                            (worksheet.Cells[row, 9].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 9].Value.ToString()))
                            )
                            {
                                throw new Exception("Se detecto un registro fuera del formato solicitado");
                            }
                            bus = new Autobus();
                            bus.Eco = int.Parse(worksheet.Cells[row, 2].Value.ToString());
                            bus.Año = int.Parse(worksheet.Cells[row, 3].Value.ToString());
                            bus.Mod = worksheet.Cells[row, 4].Value.ToString();
                            bus.Eje = worksheet.Cells[row, 5].Value.ToString();
                            bus.Sig = worksheet.Cells[row, 6].Value.ToString();
                            bus.Marca = worksheet.Cells[row, 7].Value.ToString();
                            bus.Base = worksheet.Cells[row, 8].Value.ToString();
                            bus.Rol = worksheet.Cells[row, 9].Value.ToString();
                            listBusNew.Add(bus);


                        }

                        col = 12;
                        for (int row = 3; worksheet.Cells[row, col].Value != null; row++)
                        {
                            if (//(worksheet.Cells[row, 1].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 1].Value.ToString())) ||
                            (worksheet.Cells[row, 12].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 12].Value.ToString())) ||
                            (worksheet.Cells[row, 13].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 13].Value.ToString())) ||
                            (worksheet.Cells[row, 14].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 14].Value.ToString())) ||
                            (worksheet.Cells[row, 15].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 15].Value.ToString())) ||
                            (worksheet.Cells[row, 16].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 16].Value.ToString())) ||
                            (worksheet.Cells[row, 17].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 17].Value.ToString())) ||
                            (worksheet.Cells[row, 18].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 18].Value.ToString())) ||
                            (worksheet.Cells[row, 19].Value == null || string.IsNullOrEmpty(worksheet.Cells[row, 19].Value.ToString()))
                            )
                            {
                                throw new Exception("Se detecto un registro fuera del formato solicitado");
                            }
                            bus = new Autobus();
                            bus.Eco = int.Parse(worksheet.Cells[row, 12].Value.ToString());
                            bus.Año = int.Parse(worksheet.Cells[row, 13].Value.ToString());
                            bus.Mod = worksheet.Cells[row, 14].Value.ToString();
                            bus.Eje = worksheet.Cells[row, 15].Value.ToString();
                            bus.Sig = worksheet.Cells[row, 16].Value.ToString();
                            bus.Marca = worksheet.Cells[row, 17].Value.ToString();
                            bus.Base = worksheet.Cells[row, 18].Value.ToString();
                            bus.Rol = worksheet.Cells[row, 19].Value.ToString();
                            listBusOld.Add(bus);
                        }

                        Autobus busAplicar = new Autobus();
                        for(int i=0; i<listBusNew.Count; i++)
                        {
                            bool busSigDiferente = false;
                            bool busBaseDiferente = false;
                            bool busRolDiferente = false;

                            for(int j=0; j<listBusOld.Count; j++)
                            {
                                if (listBusNew[i].Eco == listBusOld[j].Eco)
                                {
                                    /*if (listBusNew[i].Sig != listBusOld[i].Sig)
                                    {
                                        busSigDiferente=true;
                                    }*/

                                    if (listBusNew[i].Base != listBusOld[j].Base)
                                    {
                                        busBaseDiferente = true;
                                    }

                                    if (listBusNew[i].Rol != listBusOld[j].Rol)
                                    {
                                        busRolDiferente = true;
                                    }

                                    break;
                                }
                            }

                            if (busBaseDiferente || busRolDiferente)
                            {
                                listBusFinalAux.Add(listBusNew[i]);   
                            }
                        }

                        foreach (var busNew in listBusNew)
                        {
                            bool busEncontrado = false;
                            foreach (var busOld in listBusOld)
                            {
                                if (busNew.Eco == busOld.Eco)
                                {
                                    busEncontrado = true;
                                    break;
                                }
                            }

                            if (!busEncontrado)
                            {
                                listBusAgregados.Add(busNew);
                            }
                        }

                        foreach (var busOld in listBusOld)
                        {
                            bool busEncontrado = false;
                            foreach (var busNew in listBusNew)
                            {
                                if (busNew.Eco == busOld.Eco)
                                {
                                    busEncontrado = true;
                                    break;
                                }
                            }

                            if (!busEncontrado)
                            {
                                listBusEliminados.Add(busOld);
                            }

                        }
                    }
                }

                using (var package = new ExcelPackage())
                {
                    

                    foreach(var bus in listBusFinalAux)
                    {
                        listBusFinal.Add(bus);
                    }
                    var worksheet = package.Workbook.Worksheets.Add("Final");
                    var worksheet2 = package.Workbook.Worksheets.Add("BusesCambiados");
                    var worksheet3 = package.Workbook.Worksheets.Add("Agregados");
                    var worksheet4 = package.Workbook.Worksheets.Add("Eliminados");
                    //Encabezados
                    worksheet.Cells[2, 2].Value = "Eco";
                    worksheet.Cells[2, 3].Value = "Año";
                    worksheet.Cells[2, 4].Value = "Mod";
                    worksheet.Cells[2, 5].Value = "Eje";
                    worksheet.Cells[2, 6].Value = "Sig";
                    worksheet.Cells[2, 7].Value = "Marca";
                    worksheet.Cells[2, 8].Value = "Base";
                    worksheet.Cells[2, 9].Value = "Rol";


                    worksheet2.Cells[2, 2].Value = "Eco";
                    worksheet2.Cells[2, 3].Value = "Año";
                    worksheet2.Cells[2, 4].Value = "Mod";
                    worksheet2.Cells[2, 5].Value = "Eje";
                    worksheet2.Cells[2, 6].Value = "Sig";
                    worksheet2.Cells[2, 7].Value = "Marca";
                    worksheet2.Cells[2, 8].Value = "Base";
                    worksheet2.Cells[2, 9].Value = "Rol";

                    worksheet3.Cells[2, 2].Value = "Eco";
                    worksheet3.Cells[2, 3].Value = "Año";
                    worksheet3.Cells[2, 4].Value = "Mod";
                    worksheet3.Cells[2, 5].Value = "Eje";
                    worksheet3.Cells[2, 6].Value = "Sig";
                    worksheet3.Cells[2, 7].Value = "Marca";
                    worksheet3.Cells[2, 8].Value = "Base";
                    worksheet3.Cells[2, 9].Value = "Rol";


                    worksheet4.Cells[2, 2].Value = "Eco";
                    worksheet4.Cells[2, 3].Value = "Año";
                    worksheet4.Cells[2, 4].Value = "Mod";
                    worksheet4.Cells[2, 5].Value = "Eje";
                    worksheet4.Cells[2, 6].Value = "Sig";
                    worksheet4.Cells[2, 7].Value = "Marca";
                    worksheet4.Cells[2, 8].Value = "Base";
                    worksheet4.Cells[2, 9].Value = "Rol";


                    foreach (var busNew in listBusNew) 
                    { 
                        bool repetidoBus=false;
                        foreach (var busAux in listBusFinalAux)
                        {
                            if(busNew.Eco == busAux.Eco)
                            {
                                repetidoBus = true;
                                break;
                            }
                        }

                        if (!repetidoBus) 
                        {
                            listBusFinal.Add(busNew);    
                        
                        }
                    
                    }

                    for (int i = 0; i < listBusFinal.Count; i++) 
                    {
                        worksheet.Cells[i + 3, 2].Value = listBusFinal[i].Eco;
                        worksheet.Cells[i + 3, 3].Value = listBusFinal[i].Año;
                        worksheet.Cells[i + 3, 4].Value = listBusFinal[i].Mod;
                        worksheet.Cells[i + 3, 5].Value = listBusFinal[i].Eje;
                        worksheet.Cells[i + 3, 6].Value = listBusFinal[i].Sig;
                        worksheet.Cells[i + 3, 7].Value = listBusFinal[i].Marca;
                        worksheet.Cells[i + 3, 8].Value = listBusFinal[i].Base;
                        worksheet.Cells[i + 3, 9].Value = listBusFinal[i].Rol;                                     
                    }

                    for (int i = 0; i < listBusFinalAux.Count; i++)
                    {
                        worksheet2.Cells[i + 3, 2].Value = listBusFinalAux[i].Eco;
                        worksheet2.Cells[i + 3, 3].Value = listBusFinalAux[i].Año;
                        worksheet2.Cells[i + 3, 4].Value = listBusFinalAux[i].Mod;
                        worksheet2.Cells[i + 3, 5].Value = listBusFinalAux[i].Eje;
                        worksheet2.Cells[i + 3, 6].Value = listBusFinalAux[i].Sig;
                        worksheet2.Cells[i + 3, 7].Value = listBusFinalAux[i].Marca;
                        worksheet2.Cells[i + 3, 8].Value = listBusFinalAux[i].Base;
                        worksheet2.Cells[i + 3, 9].Value = listBusFinalAux[i].Rol;
                    }

                    for (int i = 0; i < listBusAgregados.Count; i++)
                    {
                        worksheet3.Cells[i + 3, 2].Value = listBusAgregados[i].Eco;
                        worksheet3.Cells[i + 3, 3].Value = listBusAgregados[i].Año;
                        worksheet3.Cells[i + 3, 4].Value = listBusAgregados[i].Mod;
                        worksheet3.Cells[i + 3, 5].Value = listBusAgregados[i].Eje;
                        worksheet3.Cells[i + 3, 6].Value = listBusAgregados[i].Sig;
                        worksheet3.Cells[i + 3, 7].Value = listBusAgregados[i].Marca;
                        worksheet3.Cells[i + 3, 8].Value = listBusAgregados[i].Base;
                        worksheet3.Cells[i + 3, 9].Value = listBusAgregados[i].Rol;
                    }

                    for (int i = 0; i < listBusEliminados.Count; i++)
                    {
                        worksheet4.Cells[i + 3, 2].Value = listBusEliminados[i].Eco;
                        worksheet4.Cells[i + 3, 3].Value = listBusEliminados[i].Año;
                        worksheet4.Cells[i + 3, 4].Value = listBusEliminados[i].Mod;
                        worksheet4.Cells[i + 3, 5].Value = listBusEliminados[i].Eje;
                        worksheet4.Cells[i + 3, 6].Value = listBusEliminados[i].Sig;
                        worksheet4.Cells[i + 3, 7].Value = listBusEliminados[i].Marca;
                        worksheet4.Cells[i + 3, 8].Value = listBusEliminados[i].Base;
                        worksheet4.Cells[i + 3, 9].Value = listBusEliminados[i].Rol;
                    }

                    // Crear tabla
                    var tblFinal = worksheet.Cells[2, 2, listBusFinal.Count + 2, 9];
                    var tblCambios = worksheet2.Cells[2, 2, listBusFinalAux.Count + 2, 9];
                    var tblAgregados = worksheet3.Cells[2, 2, listBusFinalAux.Count + 2, 9];
                    var tblEliminados = worksheet4.Cells[2, 2, listBusFinalAux.Count + 2, 9];

                    var table = worksheet.Tables.Add(tblFinal, "TablaFinal");
                    var table2 = worksheet2.Tables.Add(tblCambios, "TablaCambiosBus");
                    var table3 = worksheet3.Tables.Add(tblAgregados, "TablaBusesAgreagos");
                    var table4 = worksheet4.Tables.Add(tblEliminados, "TablaBusesEliminados");

                    table.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                    table2.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                    table3.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
                    table4.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;

                    // Generar el archivo en memoria
                    var stream = new MemoryStream(package.GetAsByteArray());
                    var fileName = "GeneratedExcel.xlsx";
                    var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                    memoryaGlobal = stream;

                }
                return (new { listaEliminados = listBusEliminados, listaAgregados=listBusAgregados,listaCambios= listBusFinalAux, listaFinal=listBusFinal},
                    memoryaGlobal);
            }
            catch (Exception ex) 
            {
                return (new{ Error = true, Message = ex.Message },null);
            }
        }
    }
}
