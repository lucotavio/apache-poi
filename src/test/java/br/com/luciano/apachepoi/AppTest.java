package br.com.luciano.apachepoi;

import static org.junit.Assert.assertTrue;

import br.com.luciano.apachepoi.model.Cliente;
import br.com.luciano.apachepoi.model.exel.CriarArquivoExcel;
import org.junit.Before;
import org.junit.Test;

import java.util.ArrayList;
import java.util.List;

public class AppTest {

    private List<Cliente> clientes = new ArrayList<>();

    @Before
    public void setup(){
        clientes.add(new Cliente(1, "Karine", "karine@hotmail.com", "856956489"));
        clientes.add(new Cliente(2, "Luciano", "luciano@hotmail.com", "8985632584"));
        clientes.add(new Cliente(3, "Pedro", "pedro@hotmail.com", "987456325"));
    }
    @Test
    public void shouldAnswerWithTrue() {
        CriarArquivoExcel criarArquivoExcel = new CriarArquivoExcel();
        criarArquivoExcel.criarArquivo("clientes.xlsx", clientes);
    }
}
