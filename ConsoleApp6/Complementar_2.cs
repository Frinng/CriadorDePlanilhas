namespace ConsoleApp6;
using ClosedXML.Excel;
using System.Collections.Generic;

public class Complementar_2 {

    public static void AlterarDadosDaplanilha() {

        if (Complementar.CaminhoCOmpleto != null) {
         
            string CaminhoDaLeitura = Complementar.CaminhoCOmpleto;

            using (var LivroDeTrabalho = new XLWorkbook(CaminhoDaLeitura)) {

                var Planilha = LivroDeTrabalho.Worksheet(1);
                
                Console.WriteLine("-----------------------------------------------");
                Console.WriteLine("|                Estoque Atual                |");
                Console.WriteLine("-----------------------------------------------");
                
                Console.WriteLine(" ");
                
                Console.WriteLine("-------------------------------------------------------------------------------");
                Console.WriteLine("|  Item  |   Descrição   |   Quantidade   |    Marca/Modelo   |  Localização  |");
                Console.WriteLine("-------------------------------------------------------------------------------");
                
                Console.WriteLine(" ");

                foreach (var linha in Planilha.RowsUsed().Skip(1)) {

                    int item = linha.Cell(1).GetValue<int>();
                    string Descri = linha.Cell(2).GetValue<string>();
                    int quantidade = linha.Cell(3).GetValue<int>();
                    string marca = linha.Cell(4).GetValue<string>();
                    string local = linha.Cell(5).GetValue<string>();
                    
                    Console.WriteLine($"{item,-4} | {Descri, 20} | {marca,12} | {local}");

                }
                
            }
            
        }else {
            
            FuncoesUteis.VOCENAOCADASTROUNADA();
            
        }  

    }
}