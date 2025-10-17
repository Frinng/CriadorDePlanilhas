﻿using System;
 using ClosedXML.Excel;

 namespace ConsoleApp6;

public class program {
    static void Main() {
        string CaminhoDoBanco = "C:\\Users\\Pedro\\Documents\\My Games\\TesteInterdace\\BancoDeDados.xlsx";

        bool sair = false;
        int OPcao = 0;
        string OPcaostring;

        while (!sair) {
            Console.Clear();
            
            Console.WriteLine("-----------------------------------------------");
            Console.WriteLine("|         Sistema Chains of Caos Login        |");
            Console.WriteLine("-----------------------------------------------");
            
            Console.WriteLine(" ");
            Console.WriteLine(" ");
            
            Console.WriteLine("-----------------------------------------------");
            Console.WriteLine("| 1-Fazer Login                               |");
            Console.WriteLine("| 2-Cadastrar Novo Usuário                    |");
            Console.WriteLine("| 3-Sair                                      |");
            Console.WriteLine("-----------------------------------------------");
            OPcaostring = Console.ReadLine();

            if (int.TryParse(OPcaostring,out OPcao)) {
                
                switch (OPcao) {
                    case 1:
                        if (FazerLogin(CaminhoDoBanco)) {
                            // Se login foi bem-sucedido, continua dentro do método
                        }
                        else {
                            Console.WriteLine("Pressione qualquer tecla para tentar novamente...");
                            Console.ReadKey();
                        }
                        break;

                    case 2:
                        LeitorPlanilha.ADicionarUsuario();
                        Console.WriteLine("\nUsuário cadastrado com sucesso!");
                        Console.WriteLine("Pressione qualquer tecla para voltar...");
                        Console.ReadKey();
                        break;

                    case 3:
                        sair = true;
                        break;

                    default:
                        Console.WriteLine("Opção inválida!");
                        Console.ReadKey();
                        break;
                }
            }else {
               
                FuncoesUteis.DIGITEUMNUMEROINTEIRO();
                
           }
        }
    }
    
    static bool FazerLogin(string CaminhoDoBanco) {
        Console.Clear();
        Console.WriteLine("-----------------------------------------------");
        Console.WriteLine("|                    LOGIN                    |");
        Console.WriteLine("-----------------------------------------------");
        
        Console.WriteLine(" ");
        
        Console.Write("Usuário: ");
        string nome = Console.ReadLine();

        Console.Write("Senha: ");
        string senha = Console.ReadLine();

        using (var wb = new XLWorkbook(CaminhoDoBanco)) {
            
            var ws = wb.Worksheet(1);
            int ultimaLinha = ws.LastRowUsed().RowNumber();

            for (int i = 2; i <= ultimaLinha; i++) {
                string nomePlanilha = ws.Cell(i, 1).GetString();
                string senhaPlanilha = ws.Cell(i, 2).GetString();

                if (nomePlanilha == nome && senhaPlanilha == senha) {
                    Console.Clear();
                    Console.WriteLine($"Login realizado com sucesso!\nBem-vindo, {nome}!");
                    
                    if (nome.Equals("ADMIN", StringComparison.OrdinalIgnoreCase)) {
                        Console.WriteLine("\nVocê está logado como ADMINISTRADOR.");
                        Console.WriteLine("Pressione qualquer tecla para acessar o menu administrativo...");
                        Console.ReadKey();
                        MenuGluGLu.MenuGluglu(); 
                    }
                    else {
                        Console.WriteLine("\nVocê está logado como usuário comum.");
                        Console.WriteLine("Pressione qualquer tecla para continuar...");
                        Console.ReadKey();
                    }

                    return true;
                }
            }
        }

        Console.WriteLine("\nUsuário ou senha incorretos!");
        return false;
    }
}