package br.com.luciano.apachepoi;

import br.com.luciano.apachepoi.model.Cliente;
import br.com.luciano.apachepoi.model.exel.CriarArquivoExcel;
import java.util.ArrayList;
import java.util.List;

/**
 * Hello world!
 */
public class App 
{
    public static void main( String[] args )
    {
        List<Cliente> clientes = new ArrayList<>();
        clientes.add(new Cliente(1, "Karine", "karine@hotmail.com", "856956489"));
        clientes.add(new Cliente(2, "Luciano", "luciano@hotmail.com", "8985632584"));
        clientes.add(new Cliente(3, "Pedro", "pedro@hotmail.com", "987456325"));
        
        CriarArquivoExcel criarArquivoExcel = new CriarArquivoExcel();
        criarArquivoExcel.criarArquivo("clientes.xlsx", clientes);
    }
}
