using System;
using System.Collections.Generic;
using System.Linq;

namespace ConsoleApp6
{
    public class Emprestimo
    {
        public string NomeOperador { get; set; }
        public string NomeFerramenta { get; set; }
        public string ModeloFerramenta { get; set; }
        public DateTime DataEmprestimo { get; set; }
        public DateTime DataDevolucao { get; set; }
        public string CodigoOperador { get; set; }
        public string CodigoFerramenta { get; set; }
        public bool Devolvido { get; set; }
    }

    public class SistemaEmprestimoIntegrado
    {
        private static List<Emprestimo> emprestimos = new List<Emprestimo>();
        private static Random random = new Random();

        public static void MenuEmprestimo()
        {
            int opcao;
            do
            {
                Console.Clear();
                
                Console.WriteLine("----------------------------------------------");
                Console.WriteLine("|    Sistema de Emprestimo de Ferramentas    |");
                Console.WriteLine("----------------------------------------------");

                Console.WriteLine(" ");
                Console.WriteLine(" ");
                
                Console.WriteLine("----------------------------------------------");
                Console.WriteLine("| 1 - Empréstimo de Ferramenta               |");
                Console.WriteLine("| 2 - Devolução de Ferramenta                |");
                Console.WriteLine("| 3 - Voltar                                 |");
                Console.WriteLine("| Escolha uma opção (1, 2 ou 3)              |");
                Console.WriteLine("----------------------------------------------");

                if (int.TryParse(Console.ReadLine(), out opcao))
                {
                    switch (opcao)
                    {
                        case 1:
                            RealizarEmprestimo();
                            break;
                        case 2:
                            RealizarDevolucao();
                            break;
                        case 3:
                            Console.Clear();
                            Console.WriteLine("Voltando ao menu principal...");
                            break;
                        default:
                            Console.WriteLine("\nErro!!! Opção inválida!");
                            Console.ReadKey();
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("\nErro!!! Opção inválida!");
                    Console.ReadKey();
                }

            } while (opcao != 3);
        }

        static void RealizarEmprestimo()
        {
            Console.Clear();
            
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("|          Emprestimo de ferramenta          |");
            Console.WriteLine("----------------------------------------------");

            Emprestimo novoEmprestimo = new Emprestimo();

            Console.WriteLine("Por favor, digite o nome da pessoa que está pegando a ferramenta: ");
            novoEmprestimo.NomeOperador = Console.ReadLine();

            Console.WriteLine("Por favor, digite o nome da ferramenta: ");
            novoEmprestimo.NomeFerramenta = Console.ReadLine();

            Console.WriteLine("Por favor, digite o modelo da ferramenta: ");
            novoEmprestimo.ModeloFerramenta = Console.ReadLine();

            Console.WriteLine("Por favor, digite a data do empréstimo (dd/mm/aaaa): ");
            if (!DateTime.TryParse(Console.ReadLine(), out DateTime dataEmp))
            {
                Console.WriteLine("\nData inválida!!!");
                Console.ReadKey();
                return;
            }
            novoEmprestimo.DataEmprestimo = dataEmp;

            Console.WriteLine("Data de devolução (dd/mm/aaaa): ");
            if (!DateTime.TryParse(Console.ReadLine(), out DateTime dataDev))
            {
                Console.WriteLine("\nData inválida!!!");
                Console.ReadKey();
                return;
            }
            novoEmprestimo.DataDevolucao = dataDev;

            novoEmprestimo.CodigoOperador = GerarCodigoAleatorio(6);
            novoEmprestimo.CodigoFerramenta = GerarCodigoAleatorio(8);
            novoEmprestimo.Devolvido = false;

            emprestimos.Add(novoEmprestimo);

            Console.WriteLine("\nEmpresimo registrado com sucesso!");
            Console.WriteLine($"Código do Operador: {novoEmprestimo.CodigoOperador}");
            Console.WriteLine($"Código da Ferramenta: {novoEmprestimo.CodigoFerramenta}");
            Console.WriteLine($"Data de Devolução: {novoEmprestimo.DataDevolucao:dd/MM/yyyy}");
            Console.ReadKey();
        }

        static void RealizarDevolucao()
        {
            Console.Clear();
            
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("|          Devolução de Ferramenta           |");
            Console.WriteLine("----------------------------------------------");

            int tentativas = 3;
            bool devolvidoComSucesso = false;

            while (tentativas > 0 && !devolvidoComSucesso)
            {
                Console.WriteLine($"\nTentativa {4 - tentativas} de 3");
                Console.WriteLine("Nome do operador: ");
                string nomeOperador = Console.ReadLine();
                Console.WriteLine("Código do operador: ");
                string codigoOperador = Console.ReadLine();
                Console.WriteLine("Nome da ferramenta: ");
                string nomeFerramenta = Console.ReadLine();
                Console.WriteLine("Modelo da ferramenta: ");
                string modeloFerramenta = Console.ReadLine();
                Console.WriteLine("Código da ferramenta: ");
                string codigoFerramenta = Console.ReadLine();

                var emprestimo = emprestimos.FirstOrDefault(e =>
                    e.NomeOperador.Equals(nomeOperador, StringComparison.OrdinalIgnoreCase) &&
                    e.NomeFerramenta.Equals(nomeFerramenta, StringComparison.OrdinalIgnoreCase) &&
                    e.ModeloFerramenta.Equals(modeloFerramenta, StringComparison.OrdinalIgnoreCase) &&
                    !e.Devolvido);

                if (emprestimo != null)
                {
                    bool codigoOperadorCorreto = emprestimo.CodigoOperador == codigoOperador;
                    bool codigoFerramentaCorreto = emprestimo.CodigoFerramenta == codigoFerramenta;

                    if (codigoOperadorCorreto && codigoFerramentaCorreto)
                    {
                        emprestimo.Devolvido = true;
                        Console.WriteLine("\nDevolução efetuada com sucesso!");
                        Console.WriteLine($"Ferramenta: {emprestimo.NomeFerramenta}");
                        Console.WriteLine($"Operador: {emprestimo.NomeOperador}");
                        Console.WriteLine($"Devolvida em: {DateTime.Now:dd/MM/yyyy}");
                        devolvidoComSucesso = true;
                    }
                    else
                    {
                        tentativas--;
                        if (!codigoOperadorCorreto) Console.WriteLine("\nCódigo do operador incorreto!");
                        if (!codigoFerramentaCorreto) Console.WriteLine("Código da ferramenta incorreto!");
                        if (tentativas == 0)
                        {
                            Console.WriteLine("\nLimite de tentativas excedido!");
                            Console.WriteLine($"O operador {nomeOperador} será advertido!");
                            Console.WriteLine($"A ferramenta {nomeFerramenta} será considerada perdida!");
                        }
                    }
                }
                else
                {
                    tentativas--;
                    Console.WriteLine("\nEmpréstimo não encontrado!");
                    if (tentativas == 0)
                    {
                        Console.WriteLine("\nLimite de tentativas excedido!");
                        Console.WriteLine("Contate o administrador do sistema!");
                    }
                }

                if (!devolvidoComSucesso && tentativas > 0)
                {
                    Console.WriteLine("\nPressione qualquer tecla para tentar novamente...");
                    Console.ReadKey();
                    Console.Clear();
                    Console.WriteLine("----------------------------------------------");
                    Console.WriteLine("|          Devolução de Ferramenta           |");
                    Console.WriteLine("----------------------------------------------");
                }
            }

            Console.WriteLine("\nPressione qualquer tecla para continuar...");
            Console.ReadKey();
        }

        static string GerarCodigoAleatorio(int tamanho)
        {
            const string caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(caracteres, tamanho)
                .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
