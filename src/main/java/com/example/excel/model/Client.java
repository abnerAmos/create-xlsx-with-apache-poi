package com.example.excel.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;

@Data
@Builder
@AllArgsConstructor // Cria construtores com e sem parâmetros
public class Client {

    private Integer id;
    private String name;
    private String email;
    private String contact;
}
