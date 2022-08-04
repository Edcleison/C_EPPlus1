using C_EPPlus1.Serialization;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace C_EPPlus1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string caminhoPlanilha = @"C:\Projetos\C_EPPlus1\C_EPPlus1\Dados\PessoaEndereco.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excel = new ExcelPackage();


            var jsonPessoa = File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + @"\pessoa.json");
            var jsonEndereco = File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + @"\endereco.json");

            var pessoas = JsonConvert.DeserializeObject<List<Pessoa>>(jsonPessoa);
            var enderecos = JsonConvert.DeserializeObject<List<Endereco>>(jsonEndereco);

           
            
            var workSheet = excel.Workbook.Worksheets.Add("Planilha 1");

            workSheet.Cells[1, 1].Value = "Id";
            workSheet.Cells[1, 2].Value = "Nome";
            workSheet.Cells[1, 3].Value = "Idade";
            workSheet.Cells[1, 4].Value = "Empresa";            
            workSheet.Cells[1, 5].Value = "Logradouro";            
            workSheet.Cells[1, 6].Value = "Numero";            
            workSheet.Cells[1, 7].Value = "Bairro";            
            workSheet.Cells[1, 8].Value = "Cidade";            
            workSheet.Cells[1, 9].Value = "UF";            
            int indice = 2;

            foreach (var pessoa in pessoas)
            {
                workSheet.Cells[indice, 1].Value = pessoa.Id;
                workSheet.Cells[indice, 2].Value = pessoa.Nome;
                workSheet.Cells[indice, 3].Value = pessoa.Idade;
                workSheet.Cells[indice, 4].Value = pessoa.Empresa;                
                indice++;

            }
             indice = 2;
            foreach (var endereco in enderecos)
            {
                workSheet.Cells[indice, 5].Value = endereco.Logradouro;
                workSheet.Cells[indice, 6].Value = endereco.Numero;
                workSheet.Cells[indice, 7].Value = endereco.Bairro;
                workSheet.Cells[indice, 8].Value = endereco.Cidade;
                workSheet.Cells[indice, 9].Value = endereco.UF;
                indice++;

            }


            FileStream arquivo = File.Create(caminhoPlanilha);
            arquivo.Close();

            File.WriteAllBytes(caminhoPlanilha, excel.GetAsByteArray());
            

            int rows = workSheet.Dimension.End.Row-1;

            var queryCPFCNPJ = workSheet.Cells[1, 1, rows, 1]
                                     .Where(c => c.Value is not null)
                                     .Select(c => c.Value.ToString())
                                     .Distinct()
                                     .ToList();

            Console.Write("Planilhas geradas: ");

            excel.Dispose();




        }
    }
}
