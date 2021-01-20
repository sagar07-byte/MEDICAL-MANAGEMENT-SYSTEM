package com.mcdevelopers.notifyme;

import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;

import android.app.Notification;
import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.app.PendingIntent;
import android.content.Intent;
import android.graphics.Color;
import android.os.Build;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;

public class MainActivity extends AppCompatActivity {

    Button notify;
    EditText e;

    @RequiresApi(api = Build.VERSION_CODES.O)
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        notify = (Button) findViewById(R.id.button);
        e = (EditText) findViewById(R.id.editText);


        final NotificationManager  manager=(NotificationManager) getSystemService(NOTIFICATION_SERVICE);;
         final String CHANNEL_ONE_ID = "com.mcdevelopers.notify.me";
        final String CHANNEL_ONE_NAME = "Channel One";

// Create the channel object, using the channel ID//

        NotificationChannel notificationChannel = new NotificationChannel(CHANNEL_ONE_ID,
                CHANNEL_ONE_NAME, NotificationManager.IMPORTANCE_HIGH);
        manager.createNotificationChannel(notificationChannel);

        notify.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intent = new Intent(MainActivity.this, MainActivity.class);
                PendingIntent pending = PendingIntent.getActivity(MainActivity.this, 0, intent, 0);
                Notification noti = new Notification.Builder(MainActivity.this,CHANNEL_ONE_ID).setContentTitle("New Message").setContentText
                        (e.getText().toString()).setSmallIcon(R.mipmap.ic_launcher).setContentIntent(pending).build();
                noti.flags |= Notification.FLAG_AUTO_CANCEL;
                manager.notify(0, noti);
            }
        });

    }

    }