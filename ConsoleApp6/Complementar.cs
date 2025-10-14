﻿using System.Runtime;

namespace ConsoleApp6;
using ClosedXML.Excel;
using System.Collections.Generic;

public class Pessoa {
    
    public string Descricao { get; set; }
    public int Quantidade { get; set; }
    public string Marca_Modelo { get; set; }
    public string Localiza { get; set; }
    
    public static void pessoa(){}
    
}

public class Complementar {

    public static int Itens = 0;
    public static List<Pessoa> Pessoas = new List<Pessoa>();
    public static string NomePLanilha;

    public static void NomedaPlanilha() {

        Console.WriteLine("Digite o nome da planilha: ");
        NomePLanilha = Console.ReadLine();
        
        if (string.IsNullOrWhiteSpace(NomePLanilha)) {
            
            Console.WriteLine("Você não digitou nada!");
            
        }

    }
    
    public static void NumeroBancadas() {
        string ItensString;
        int n1;
        
        Console.WriteLine("Digite Quantos itens voce deseja adicionar: ");
        ItensString = Console.ReadLine();

        if (int.TryParse(ItensString, out n1)) {
            
            Itens = n1;
            
        }
        else {
            
            FuncoesUteis.DIGITEUMNUMEROINTEIRO();
            
        }
    }
    
    public static void CadastroParaAPlanilha(){
        int n1;

        for (int i = 0; i < Itens; i++) {
            Console.Clear();
            
            Pessoa Pessoane = new Pessoa();
            Console.WriteLine("--------------------------------------------");
            
            Console.WriteLine("Digite o nome do item: ");
            Pessoane.Descricao = Console.ReadLine();
            
            Console.WriteLine(" ");
        
            Console.WriteLine($"Digite a quantidade de {Pessoane.Descricao}");
            string QuantidadeDigito = Console.ReadLine();

            if (int.TryParse(QuantidadeDigito,out n1)) {

                Pessoane.Quantidade = n1;

            }else {
          
                FuncoesUteis.DIGITEUMNUMEROINTEIRO();
            
            }
            
            Console.WriteLine(" ");
            
            Console.WriteLine("Digite a Marca/Modelo: ");
            Pessoane.Marca_Modelo = Console.ReadLine();
            
            Console.WriteLine(" ");
            
            Console.WriteLine("Digite a localização do item: ");
            Pessoane.Localiza = Console.ReadLine();
            
            Pessoas.Add(Pessoane);
            
            Console.WriteLine("--------------------------------------------");
        }
    }

    public static void CadastroPlanilha() {

        int Contador = 0;
        
        string CaminhoDaPlanilhakkkslkem = "C:\\Users\\Pedro\\Documents\\My Games\\TesteInterdace";

        string CaminhoCOmpleto = Path.Combine(CaminhoDaPlanilhakkkslkem, NomePLanilha + ".xlsx");
        
        using (var LivroDeTrabalho = new XLWorkbook()) {

            var Planilha = LivroDeTrabalho.Worksheets.Add(NomePLanilha);


            Planilha.Cell(1, 1).Value = "Item";
            Planilha.Cell(1, 2).Value = "Descrição";
            Planilha.Cell(1, 3).Value = "Quantidade";
            Planilha.Cell(1, 4).Value = "Marca/Modelo";
            Planilha.Cell(1, 5).Value = "localização";
            
            for (int j = 0; j < Pessoas.Count && j < Pessoas.Count; j++) {

                Contador++;
                
                Planilha.Cell(2 + j, 1).Value = Contador;
                Planilha.Cell(2 + j, 2).Value = Pessoas[j].Descricao;
                Planilha.Cell(2 + j, 3).Value = Pessoas[j].Quantidade;
                Planilha.Cell(2 + j, 4).Value = Pessoas[j].Marca_Modelo;
                Planilha.Cell(2 + j, 5).Value = Pessoas[j].Localiza;
                
            }
            
            LivroDeTrabalho.SaveAs(CaminhoCOmpleto);
        }
    }
}