﻿namespace ConsoleApp6;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

public class Complementar_2 {
    
    public static string NumeroItem;
    public static int N1 = 0;
    public static int numeroitem = 0;
    public static int quantidaAtual;
    public static string CaminhoDoEstoque; 
    public static void LerPlanilhas() {

        CaminhoDoEstoque = "C:\\Users\\Pedro\\Documents\\My Games\\TesteInterdace\\PLanilhaSimuladoEstoque.xlsx";
        
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
                    
                    Console.WriteLine($"| {item,-4}|  {Descri, 20}   |   {quantidade}   |  {marca,12}  | {local}  |");

                }
                
                Console.WriteLine(" ");
                        
                Console.WriteLine("-------------------------------------------------------------------------------");
                Console.WriteLine("|                    Pressione qualquer tecla para continuar                  |");
                Console.WriteLine("-------------------------------------------------------------------------------");

                Console.ReadKey();
        }
        
    }
    
    public static void ADICIONARITEM() {

        string Numerostring;
        int numer = 0;
        int quantidade = 0;
        
        if (CaminhoDoEstoque == null) {
            
            FuncoesUteis.SistemaForaDOaR();
            
            return;
        }

        using (var LivroTrabalho = new XLWorkbook(CaminhoDoEstoque)) {

            var Planilha1 = LivroTrabalho.Worksheet(1);

            int ultimalinha = Planilha1.LastRowUsed().RowNumber() + 1;
            int novoNumero = ultimalinha - 1;
            
            Console.WriteLine("Digite o nome do item: ");
            string descricao = Console.ReadLine();
            
            Console.WriteLine($"Digite a quantidade de {descricao}: ");
            Numerostring = Console.ReadLine();

            if (int.TryParse(Numerostring, out numer)) {

                quantidade = numer;

            }
            else {
               
                FuncoesUteis.DIGITEUMNUMEROINTEIRO();
                
            }
            
            Console.WriteLine("Digite a Marca/Modelo: ");
            string Marca = Console.ReadLine();
            
            Console.WriteLine($"Digite onde o {descricao} sera armazenado: ");
            string locali = Console.ReadLine();
            
            Planilha1.Cell(ultimalinha,1).Value = novoNumero;
            Planilha1.Cell(ultimalinha, 2).Value = descricao;
            Planilha1.Cell(ultimalinha, 3).Value = quantidade;
            Planilha1.Cell(ultimalinha, 4).Value = Marca;
            Planilha1.Cell(ultimalinha, 5).Value = locali;
            
            LivroTrabalho.Save();
            
        }
        
    }
    
    public static void REMOVERUMALINHACOMPLETA() {
        string numerostring;
        int numero = 0;
        int NumeroITemRemove = 0;
        
        string CaminhoRemove = "C:\\Users\\Pedro\\Documents\\My Games\\TesteInterdace\\PLanilhaSimuladoEstoque.xlsx";
        
        if (CaminhoRemove == null) {
            
            FuncoesUteis.SistemaForaDOaR();
            
            return;
            
        }

        using (var BookDeTrampo = new XLWorkbook(CaminhoRemove)) {
            
            var Planilha2 = BookDeTrampo.Worksheet(1);
            
            Console.WriteLine("Digite o numero do item que você deseja remover: ");
            numerostring = Console.ReadLine();

            if (int.TryParse(numerostring, out numero)) {

                NumeroITemRemove = numero;

            }
            else {
                
                FuncoesUteis.DIGITEUMNUMEROINTEIRO();
                
            }

            var linheAchada = Planilha2.RowsUsed().Skip(1).FirstOrDefault(r => r.Cell(1).GetValue<int>() == NumeroITemRemove);

            if (linheAchada == null) {
                
                FuncoesUteis.ITEMNAOENCONTRADO();
                
                return;
                
            }

            linheAchada.Delete();
            
            BookDeTrampo.Save();
            
        }
        
    }

    public static void AdicionarProduto() {
        
        int linhaencontrada = -1;
        string nome;
        string adicionarstring;
        int adicionar = 0;
        
        string CaminhoAdicionarProduto = "C:\\Users\\Pedro\\Documents\\My Games\\TesteInterdace\\PLanilhaSimuladoEstoque.xlsx";
        
        if (CaminhoAdicionarProduto == null) {
            
            FuncoesUteis.SistemaForaDOaR();
            
            return;
        }

        using (var TrabalhoLivro = new XLWorkbook(CaminhoAdicionarProduto)) {

            var PlanilhaAdicionar = TrabalhoLivro.Worksheet(1);
            
            Console.WriteLine("Digite o numero do item que voce deseja adicionar: ");
            NumeroItem = Console.ReadLine();
        
            if (int.TryParse(NumeroItem, out N1)) {

                numeroitem = N1;

            }else {
                    
                FuncoesUteis.DIGITEUMNUMEROINTEIRO();
                    
            }
        
            foreach (var linha in PlanilhaAdicionar.RowsUsed().Skip(1)) {
                if (linha.Cell(1).GetValue<int>() == numeroitem) {

                    linhaencontrada = linha.RowNumber();
                        
                    break;

                }
            }
            if (linhaencontrada == -1) {

                FuncoesUteis.ITEMNAOENCONTRADO();
                    
                return;
            }
            
            var celulaquantida = PlanilhaAdicionar.Cell(linhaencontrada, 3);
            quantidaAtual = celulaquantida.GetValue<int>();
            nome = PlanilhaAdicionar.Cell(linhaencontrada, 2).GetValue<string>();
                
            Console.WriteLine($"Quantidade Atual de {nome}: {quantidaAtual}");
                
            Console.WriteLine(" ");
            
            Console.WriteLine("Digite quanto voce deseja adicionar: ");
            adicionarstring = Console.ReadLine();

            if (int.TryParse(adicionarstring, out adicionar)) {

                quantidaAtual += adicionar;
                                
                celulaquantida.Value = quantidaAtual;
                                    
                TrabalhoLivro.Save();
            }
            else {
                                
                FuncoesUteis.DIGITEUMNUMEROINTEIRO();
                                
            }
        }
        
    }

    public static void RemoverProduto() {
        
        int linhaencontrada = -1;
        string nome;
        string retirarstring;
        string NumerItem;
        int Numer1 = 0;
        int numerItem = 0;
        int QuantidAtual = 0;
        int retirar = 0;
        
        
        string CaminhoRemoverProduto = "C:\\Users\\Pedro\\Documents\\My Games\\TesteInterdace\\PLanilhaSimuladoEstoque.xlsx";
        
        if (CaminhoRemoverProduto == null) {
            
            FuncoesUteis.SistemaForaDOaR();
            
            return;
        }

        using (var TrabalhoLivroRemover = new XLWorkbook(CaminhoRemoverProduto)) {

            var PlanilhaRemover = TrabalhoLivroRemover.Worksheet(1);
            
            Console.WriteLine("Digite o numero do item que voce deseja Remover: ");
            NumerItem = Console.ReadLine();
        
            if (int.TryParse(NumerItem, out Numer1)) {

                numerItem = Numer1;

            }else {
                    
                FuncoesUteis.DIGITEUMNUMEROINTEIRO();
                    
            }
            foreach (var linha in PlanilhaRemover.RowsUsed().Skip(1)) {
                if (linha.Cell(1).GetValue<int>() == numerItem) {

                    linhaencontrada = linha.RowNumber();
                        
                    break;

                }
            }
            if (linhaencontrada == -1) {

                FuncoesUteis.ITEMNAOENCONTRADO();
                    
                return;
            }
            var celulaquantida = PlanilhaRemover.Cell(linhaencontrada, 3);
            QuantidAtual = celulaquantida.GetValue<int>();
            nome = PlanilhaRemover.Cell(linhaencontrada, 2).GetValue<string>();
            
            Console.WriteLine($"Quantidade Atual de {nome}: {QuantidAtual}");
                
            Console.WriteLine(" ");
            
            Console.WriteLine("Digite quanto voce deseja retirar: ");
            retirarstring = Console.ReadLine();

            if (int.TryParse(retirarstring, out retirar)) {

                if (QuantidAtual >= retirar) {
                                    
                    QuantidAtual -= retirar;
                                    
                    celulaquantida.Value = QuantidAtual;
                                    
                    TrabalhoLivroRemover.Save();

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
            
        }
    }
}