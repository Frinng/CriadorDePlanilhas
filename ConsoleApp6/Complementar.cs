﻿using System.Runtime;

namespace ConsoleApp6;
using ClosedXML.Excel;
using System.Collections.Generic;

public class Pessoa {
    
    public string Nome { get; set; }
    public int Idade { get; set; }
    public string Cidade { get; set; }
    
    public static void pessoa(){}
    
}

public class Complementar
{

    public static int Bancadas = 1;
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
        string BancadasString;
        int n1;
        
        Console.WriteLine("Digite Quantas Bancadas vc quer adicionar: ");
        BancadasString = Console.ReadLine();

        if (int.TryParse(BancadasString, out n1)) {
            Bancadas = n1;
        }
        else {
            
            FuncoesUteis.DIGITEUMNUMEROINTEIRO();
            
        }
    }
    
    public static void CadastroParaAPlanilha(){
        int n1;

        for (int i = 0; i <= 1; i++) {
            Pessoa Pessoane = new Pessoa();
            Console.WriteLine("--------------------------------------------");
            
            Console.WriteLine("Digite um nome: ");
            Pessoane.Nome = Console.ReadLine();
        
            Console.WriteLine($"Digite a idade do {Pessoane.Nome}");
            string IdadeDigito = Console.ReadLine();

            if (int.TryParse(IdadeDigito,out n1)) {

                Pessoane.Idade = n1;

            }else {
          
                FuncoesUteis.DIGITEUMNUMEROINTEIRO();
            
            }
            
            Console.WriteLine($"Digite o nome da cidade em que o {Pessoane.Nome} vive: ");
            Pessoane.Cidade = Console.ReadLine();
            
            Pessoas.Add(Pessoane);
            
            Console.WriteLine("  ");
        }
    }

    public static void CadastroPlanilha() {
        
        string CaminhoDaPlanilhakkkslkem = "C:\\Users\\Pedro\\Documents\\My Games\\TesteInterdace";

        string CaminhoCOmpleto = Path.Combine(CaminhoDaPlanilhakkkslkem, NomePLanilha + ".xlsx");
        
        using (var LivroDeTrabalho = new XLWorkbook()) {

            var Planilha = LivroDeTrabalho.Worksheets.Add(NomePLanilha);


            Planilha.Cell(1, 1).Value = "Nome";
            Planilha.Cell(1, 2).Value = "Idade";
            Planilha.Cell(1, 3).Value = "Cidade";
            
            for (int j = 0; j < Pessoas.Count && j < Pessoas.Count; j++) {
                
                Planilha.Cell(2 + j, 1).Value = Pessoas[j].Nome;
                Planilha.Cell(2 + j, 2).Value = Pessoas[j].Idade;
                Planilha.Cell(2 + j, 3).Value = Pessoas[j].Cidade;
                
            }
            
            LivroDeTrabalho.SaveAs(CaminhoCOmpleto);
        }
    }
}