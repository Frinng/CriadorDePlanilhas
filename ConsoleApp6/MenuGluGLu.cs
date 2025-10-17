namespace ConsoleApp6;

public class MenuGluGLu {
    
    public static void MenuGluglu() {
        var Continuar = true;
        string Escolhastring;
        int Escolha = 0;

        do {
            Console.Clear();
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("|            Sistema de Compras              |");
            Console.WriteLine("----------------------------------------------");
            
            Console.WriteLine(" ");
            Console.WriteLine(" ");
            
            Console.WriteLine("----------------------------------------------");
            Console.WriteLine("| 1 - Adicionar um item                      |");
            Console.WriteLine("| 2 - Remover um item                        |");
            Console.WriteLine("| 3 - Adicionar Produto                      |");
            Console.WriteLine("| 4 - Remover Produto                        |");
            Console.WriteLine("| 5 - Ver Planilha                           |");
            Console.WriteLine("| 6 - Sair                                   |");
            Console.WriteLine("----------------------------------------------");
            Escolhastring = Console.ReadLine();

            if (int.TryParse(Escolhastring, out Escolha)) {
                
                switch (Escolha) {
                    case 1:
                        Console.Clear();
                        
                        Complementar_2.ADICIONARITEM();
                        
                        FuncoesUteis.CLIQUEAQUIPARACONTINUAR();
                        break;
                    case 2:
                        Console.Clear();
                        
                        Complementar_2.REMOVERUMALINHACOMPLETA();
                        
                        FuncoesUteis.CLIQUEAQUIPARACONTINUAR();
                        break;
                    case 3:
                        Console.Clear();
                        
                        Complementar_2.AdicionarProduto();
                        
                        FuncoesUteis.CLIQUEAQUIPARACONTINUAR();
                        break;
                    case 4:
                        Console.Clear();
                        
                        Complementar_2.RemoverProduto();
                        
                        FuncoesUteis.CLIQUEAQUIPARACONTINUAR();
                        break;
                    case 5:
                        Console.Clear();
                        
                        Complementar_2.LerPlanilhas();
                        
                        break;
                    case 6:
                        
                        FuncoesUteis.SAINDODOSISTEMA();
                        
                        Continuar = false;
                        break;
                    default:
                        
                        
                        FuncoesUteis.DIGITEUMAOPCAOVALIDA();
                        break;
                }
            }
            else {
                
                FuncoesUteis.DIGITEUMAOPCAOVALIDA();
                
            }
        } while (Continuar);
    }
}