﻿using System;
namespace ConsoleApp6;

public class program {
    public static void Main() {
        
        var Continuar = true;
        
        string Escolhastring;
        int Escolha = 0;
                
        do {
                    
            //Limpa A Tela
            Console.Clear();
                         
            //Menu Principal
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("|            Sistema de Compras              |");
            Console.WriteLine("----------------------------------------------");
                         
            Console.WriteLine(" ");
            Console.WriteLine(" ");
                         
            //opcoes
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("| 1-Adicionar um item                        |");
            Console.WriteLine("| 2-Ver Planilha                             |");
            Console.WriteLine("| 3-Sair                                     |");
            Console.WriteLine("----------------------------------------------");
            Escolhastring = Console.ReadLine();
        
            if (int.TryParse(Escolhastring, out Escolha)) {
                switch (Escolha) {
                    case 1:
                        Console.Clear();

                        ProgramaCompleto.Completin();
                        
                        FuncoesUteis.CLIQUEAQUIPARACONTINUAR();
                                
                        break;
                    case 2:
                        
                        Complementar_2.LerPlanilhas();
                        
                        break;
                    case 3:
                                
                        FuncoesUteis.SAINDODOSISTEMA();
        
                        Continuar = false;
                                
                        break;
                    default:
                        FuncoesUteis.DIGITEUMAOPCAOVALIDA();
                                
                        break;
                            
                }
            }else {
                        
                FuncoesUteis.DIGITEUMAOPCAOVALIDA();
                        
            }
        } while (Continuar);
    }
}