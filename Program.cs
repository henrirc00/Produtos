using System;
using NetOffice.ExcelApi;
using Produtos.Classes;


namespace Produtos
{
    class Program
    {
        static void Main(string[] args)
        {
           Produto pro = new Produto();

           Categoria cat = new Categoria();
           cat.nome = "Bebida";
           cat.descricao = "lanches bom";

           Fornecedor dor = new Fornecedor();
           dor.nomeFantasia = "burguerr";
           dor.razaoSocial = "HR";

           pro.nome = "Guaraviton";
           pro.descricao  = "Bebida";          
           pro.preco = 3; 
           pro.quantidade = 1;
           pro.fornecedor = dor;
           pro.categoria = cat;
           Console.WriteLine(pro.cadastrar());

           
           
        


        }
    }
}
