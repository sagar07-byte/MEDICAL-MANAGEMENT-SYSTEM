package com.example.mycalculator;

import androidx.appcompat.app.AppCompatActivity;

import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;

public class MainActivity extends AppCompatActivity {

    EditText num1, num2;
    TextView result;
    Button   add, sub, mult, div, clear;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        num1=findViewById(R.id.num1);
        num2=findViewById(R.id.num2);
        result=findViewById(R.id.result);
        add=findViewById(R.id.btnAdd);
        sub=findViewById(R.id.btnSub);
        mult=findViewById(R.id.btnMul);
        div=findViewById(R.id.btnDiv);
        clear=findViewById(R.id.btnClear);

        clear.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                num1.setText("");
                num2.setText("");
                result.setText("");
            }
        });

    }

    public void addition(View view){

        int one= Integer.parseInt(num1.getText().toString());
        int two= Integer.parseInt(num2.getText().toString());
        result.setText(String.valueOf(one+two));
    }

    public void subtraction(View view){

        int one= Integer.parseInt(num1.getText().toString());
        int two= Integer.parseInt(num2.getText().toString());
        result.setText(String.valueOf(one-two));
    }

    public void multiplication(View view){

        int one= Integer.parseInt(num1.getText().toString());
        int two= Integer.parseInt(num2.getText().toString());
        result.setText(String.valueOf(one*two));
    }
    public void division(View view){

        int one= Integer.parseInt(num1.getText().toString());
        int two= Integer.parseInt(num2.getText().toString());
        result.setText(String.valueOf(one/two));
    }



}