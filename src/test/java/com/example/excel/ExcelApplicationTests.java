package com.example.excel;

import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.List;

@SpringBootTest
class ExcelApplicationTests {

    @Test
    void contextLoads() {

//        List<? extends People> people = getExtend();
//        Man man = new Man();
//        people.add(man);
//        People people1 = people.get(1);

        List<? super People> aSuper = getSuper();
        aSuper.add(new Man());
        aSuper.add(new People());

    }



    public List<? extends People> getExtend(){
        List<Man> list = new ArrayList<>();

        return list;
    }

    public List<? super People> getSuper(){
        List<People> list = new ArrayList<>();

        return list;
    }
}
