﻿namespace ConsoleApp6;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

public class Complementar_2 {
    
    public static string NumeroItem;
    public static int N1 = 0;
    public static int numeroitem = 0;
    public static int quantidaAtual;
    public static void LerPlanilhas()
    {

        string CaminhoDoEstoque =  "C:\\Users\\aluno_iot\\Documents\\Pedro\\ProjetoTestePlanilha\\Simulacao Estoque.xlsx";
        
        using (var LivroDeTrabalho = new XLWorkbook(CaminhoDoEstoque)) {

                var Planilha = LivroDeTrabalho.Worksheet(1);
                
                Console.WriteLine("-------------------------------------------------------------------------------");
                Console.WriteLine("|                                Estoque Atual                                |");
                Console.WriteLine("-------------------------------------------------------------------------------");
                
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
                    
                    Console.WriteLine($"|  {item,-4}  | {Descri, 20}  |  {quantidade}  |  {marca,12}  | {local}  |");

                }
                
                Console.WriteLine(" ");
                        
                Console.WriteLine("-------------------------------------------------------------------------------");
                Console.WriteLine("|                    Pressione qualquer tecla para continuar                  |");
                Console.WriteLine("-------------------------------------------------------------------------------");

                Console.ReadKey();
            }
        
    }

    public static void NumeroqueVoceDesejaAlterar() {
        
        Console.WriteLine("Digite o numero do item que voce deseja alterar: ");
        NumeroItem = Console.ReadLine();
        
        if (int.TryParse(NumeroItem, out N1)) {

            numeroitem = N1;

        }else {
                    
            FuncoesUteis.DIGITEUMNUMEROINTEIRO();
                    
        }
        
    }
    
    public static void AlterarDadosDaplanilha() {
        if (Complementar.CaminhoCOmpleto != null) {
            
            string CaminhoDaAlteração = Complementar.CaminhoCOmpleto;

            using (var LivroTrabalhokkkOloco = new XLWorkbook(CaminhoDaAlteração)) {

                var Planilha = LivroTrabalhokkkOloco.Worksheet(1);

                int linhaencontrada = -1;
                string nome;
                string adicionarstring;
                int adicionar = 0;
                string retirarstring;
                int retirar = 0;
                string Opcaostring;
                int Opcao = 0;

                NumeroqueVoceDesejaAlterar();

                foreach (var linha in Planilha.RowsUsed().Skip(1)) {
                    if (linha.Cell(1).GetValue<int>() == numeroitem) {

                        linhaencontrada = linha.RowNumber();
                        
                        break;

                    }
                }

                if (linhaencontrada == -1) {

                    FuncoesUteis.ITEMNAOENCONTRADO();
                    
                    return;
                }

                var celulaquantida = Planilha.Cell(linhaencontrada, 3);
                quantidaAtual = celulaquantida.GetValue<int>();
                nome = Planilha.Cell(linhaencontrada, 2).GetValue<string>();
                
                Console.WriteLine($"Quantidade Atual de {nome}: {quantidaAtual}");
                
                Console.WriteLine(" ");
                
                Console.WriteLine("----------------------------------------------");
                Console.WriteLine("|                você deseja                 |");
                Console.WriteLine("----------------------------------------------");
                
                Console.WriteLine("----------------------------------------------");
                Console.WriteLine("| 1-Adicionar                                |");
                Console.WriteLine("| 2-Retirar                                  |");
                Console.WriteLine("----------------------------------------------");
                Opcaostring =Console.ReadLine();

                if (int.TryParse(Opcaostring, out Opcao)) {

                    switch (Opcao) {
                        
                        case 1:
                            Console.Clear();
                            
                            Console.WriteLine("Digite quanto voce deseja adicionar: ");
                            adicionarstring = Console.ReadLine();

                            if (int.TryParse(adicionarstring, out adicionar)) {

                                quantidaAtual += adicionar;
                                
                                celulaquantida.Value = quantidaAtual;
                                    
                                LivroTrabalhokkkOloco.Save();
                            }
                            else {
                                
                                FuncoesUteis.DIGITEUMNUMEROINTEIRO();
                                
                            }
                            
                            break;
                        case 2:
                            Console.Clear();
                            Console.WriteLine("Digite quanto voce deseja retirar: ");
                            retirarstring = Console.ReadLine();

                            if (int.TryParse(retirarstring, out retirar)) {

                                if (quantidaAtual >= retirar) {
                                    
                                    quantidaAtual -= retirar;
                                    
                                    celulaquantida.Value = quantidaAtual;
                                    
                                    LivroTrabalhokkkOloco.Save();

                                    Console.WriteLine("\nAlteração salva com sucesso!");
                                    
                                }else {
                                    
                                     
                                    Console.WriteLine("------------------------------------------------------------------");
                                    Console.WriteLine("|    Quantidade já está em zero, não é possível retirar mais     |");
                                    Console.WriteLine("------------------------------------------------------------------");
                                        
                                }

                            }
                            else {
                                
                                FuncoesUteis.DIGITEUMNUMEROINTEIRO();
                                
                            }
                            
                            break;
                        default:
                            
                            FuncoesUteis.DIGITEUMAOPCAOVALIDA();
                            
                            break;
                        
                    }
                    
                }else {
                    
                    FuncoesUteis.DIGITEUMAOPCAOVALIDA();
                    
                }
            }
        }else {
              
              FuncoesUteis.VOCENAOCADASTROUNADA();
              
        }
    }
}