
using System.IO;
using NetOffice.ExcelApi;


namespace Produtos.Classes
{
    public class Produto
    {
    public string nome;
    public string descricao;
    public double preco;
    public double quantidade;
    public Categoria categoria;
    public Fornecedor fornecedor;


    public string cadastrar(){
        Application ex = new Application();
        FileInfo arquivo = new FileInfo(@"c:\Henrique\produto.xlsx");
        if (arquivo.Exists){
            ex.Visible = true;
            ex.Workbooks.Open(@"c:\Henrique\produto.xlsx");

            for(int y=1; y <=80; y++)
            {

                if(ex.Range("A"+y).Value == null){
                    ex.Range("a"+y).Value = nome;
                    ex.Range("b"+y).Value = descricao;
                    ex.Range("c"+y).Value = preco;
                    ex.Range("d"+y).Value = quantidade;
                    ex.Range("e"+y).Value = categoria.nome;
                    ex.Range("f"+y).Value = fornecedor.razaoSocial;

                    break;
                    }

                }
                ex.ActiveWorkbook.Save();
                ex.Quit();
             }
            else{
                ex.Visible = true;
                ex.Workbooks.Add();

                ex.Range("a1").Value = "Nome do produto";
                ex.Range("b1").Value = "Comida legal";
                ex.Range("c1").Value = "Preço";
                ex.Range("d1").Value = "Quantidade";
                ex.Range("e1").Value = "Categoria";
                ex.Range("f1").Value = "Fornecedor";

                ex.Range("a1:j1").Font.Name = "Tahoma";
                ex.Range("a1:j1").Font.Bold = true;
                ex.Range("a1:j1").Font.Size = 15; 

                                 
                ex.Range("a2").Value = nome;
                ex.Range("b2").Value = descricao;
                ex.Range("c2").Value = preco;
                ex.Range("d2").Value = quantidade;
                ex.Range("e2").Value = categoria.nome;
                ex.Range("f2").Value = fornecedor.razaoSocial;

                ex.ActiveWorkbook.SaveAs(@"c:\Henrique\produto.xlsx");
                ex.Quit();




            }


            return "Produto Cadastrado com sucesso";
            
        }

        
    

        
            
        
    }

    }
