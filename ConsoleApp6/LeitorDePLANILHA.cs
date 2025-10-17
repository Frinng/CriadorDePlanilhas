using ClosedXML.Excel;

namespace ConsoleApp6;

public class LeitorPlanilha
{
    public static void ADicionarUsuario()
    {
        string caminhoLogin = "C:\\Users\\Aluno_Analytics.CENTRO40\\Documents\\Pedro\\Planilhas\\BancoDeDados.xlsx";

        Console.Clear();
        
        Console.WriteLine("-----------------------------------------------");
        Console.WriteLine("|            CADASTRAR NOVO USUÁRIO           |");
        Console.WriteLine("-----------------------------------------------");
        
        Console.WriteLine(" ");
        
        Console.Write("Digite um nome para seu usuário: ");
        string nome = Console.ReadLine();

        Console.Write("Crie sua senha: ");
        string senha = Console.ReadLine();

        using (var LivroDeTrabalho = new XLWorkbook(caminhoLogin))
        {
            var PlanilhaLogin = LivroDeTrabalho.Worksheet("Usuarios");
            var ultimaCelulaUsada = PlanilhaLogin.LastRowUsed();
            int ultimaLinha = 1;

            if (ultimaCelulaUsada != null)
                ultimaLinha = ultimaCelulaUsada.RowNumber() + 1;
            
            for (int i = 2; i < ultimaLinha; i++) {
                if (PlanilhaLogin.Cell(i, 1).GetString() == nome) {
                    Console.WriteLine("Usuário já existe! Tente outro nome.");
                    return;
                }
            }

            PlanilhaLogin.Cell(ultimaLinha, 1).Value = nome;
            PlanilhaLogin.Cell(ultimaLinha, 2).Value = senha;

            LivroDeTrabalho.Save();
        }
    }
}