package com.example.hospitalrecord_v1;

import androidx.appcompat.app.AppCompatActivity;

import android.app.DatePickerDialog;
import android.content.Context;
import android.content.SharedPreferences;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.Button;
import android.widget.DatePicker;
import android.widget.EditText;
import android.widget.TextView;

import java.util.Calendar;

public class NameDateActivity extends AppCompatActivity {
    private String TAG = "NameDateActivity"; 
    private TextView txt_nameDate_date;
    private DatePickerDialog datePickerDialog;
    private Button btn_nameDate_send;
    private EditText edit_nameDate_name;
    private String name,date;
    private String timeFormat;
    SharedPreferences sharedPref;
    SharedPreferences.Editor editor;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_name_date);
        initView();
    }
    
    private void initView(){
        timeFormat = "";
        initUserData();
        edit_nameDate_name = (EditText)findViewById(R.id.edit_nameDate_name);
        txt_nameDate_date = (TextView) findViewById(R.id.txt_nameDate_date);
        btn_nameDate_send = (Button) findViewById(R.id.btn_nameDate_send);
        txt_nameDate_date.setOnClickListener(onClick);
        btn_nameDate_send.setOnClickListener(onClick);

        edit_nameDate_name.setText(name);
        if(date.equals("")){
            txt_nameDate_date.setText("未輸入");
        }else {
            txt_nameDate_date.setText(date);
        }
        
    }
    
    private View.OnClickListener onClick = new View.OnClickListener() {
        @Override
        public void onClick(View view) {
            if (view.getId() == R.id.txt_nameDate_date){
                Log.d(TAG, "onClick: txt_nameDate_date");
                Calendar c = Calendar.getInstance();
                datePickerDialog = new DatePickerDialog(NameDateActivity.this, new DatePickerDialog.OnDateSetListener() {
                    @Override
                    public void onDateSet(DatePicker datePicker, int year, int month, int day) {
                        timeFormat = year +"-"+ (month+1) +"-"+ day;
                        Log.d(TAG, "onDateSet: "+timeFormat);
                        txt_nameDate_date.setText(timeFormat);
                    }
                },c.get(Calendar.YEAR),c.get(Calendar.MONTH),c.get(Calendar.DAY_OF_MONTH));
                datePickerDialog.show();
            }else if (view.getId() == R.id.btn_nameDate_send){
                Log.d(TAG, "onClick: btn_nameDate_send");
                setData();
                finish();
            }
            
        }
    };

    private void initUserData(){
        sharedPref = getSharedPreferences("data_storage",Context.MODE_PRIVATE);
        editor = sharedPref.edit();
        name = sharedPref.getString("name","");
        date = sharedPref.getString("date","");
        Log.d(TAG, "initUserData: name = "+name);
        Log.d(TAG, "initUserData: date = "+date);
    }

    private void setData(){
        name = edit_nameDate_name.getText().toString();
        if(!name.equals("")){
            Log.d(TAG, "setData: setName");
            Log.d(TAG, "setData: name = "+name);
            editor.putString("name", name);
            editor.apply();
        }
        if (!timeFormat.equals("")){
            Log.d(TAG, "setData: setDate");
            editor.putString("date", timeFormat);
            editor.apply();
        }
    }
}