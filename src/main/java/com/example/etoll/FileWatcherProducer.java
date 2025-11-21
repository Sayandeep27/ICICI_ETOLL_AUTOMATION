package com.example.etoll;

import org.apache.kafka.clients.producer.KafkaProducer;
import org.apache.kafka.clients.producer.ProducerRecord;

import java.io.IOException;
import java.nio.file.*;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

public class FileWatcherProducer implements Runnable {

    private final Path root = Paths.get("dsr_reports");
    private final KafkaProducer<String,String> producer;
    private final String topic = "dsr_topic";

    public FileWatcherProducer() throws IOException {
        Properties props = KafkaConfig.getProducerProps();
        this.producer = new KafkaProducer<>(props);
    }

    @Override
    public void run() {
        processExistingFiles();   // NEW: process all existing folders first
        watchForNewFiles();       // then watch for new folders
    }

    private void processExistingFiles() {
        try {
            System.out.println("[Producer] Scanning existing folders...");

            if (!Files.exists(root)) return;

            Files.list(root).filter(Files::isDirectory).forEach(folder -> {
                Path file = folder.resolve("dsr_report.xlsx");
                if (Files.exists(file)) {
                    sendToKafka(folder, "dsr_report.xlsx");
                }
            });

        } catch (Exception e) {
            System.out.println("[Producer] ERROR reading existing folders: " + e.getMessage());
        }
    }

    private void watchForNewFiles() {
        try {
            WatchService watch = FileSystems.getDefault().newWatchService();
            root.register(watch, StandardWatchEventKinds.ENTRY_CREATE);

            System.out.println("[Producer] Watching -> " + root);

            while (true) {
                WatchKey key = watch.poll(2, TimeUnit.SECONDS);
                if (key == null) continue;

                for (WatchEvent<?> event : key.pollEvents()) {
                    if (event.kind() == StandardWatchEventKinds.ENTRY_CREATE) {
                        Path folder = root.resolve((Path) event.context());
                        Path file = folder.resolve("dsr_report.xlsx");

                        if (Files.exists(file)) {
                            sendToKafka(folder, "dsr_report.xlsx");
                        }
                    }
                }
                key.reset();
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void sendToKafka(Path folder, String fileName) {
        String f = folder.toString().replace("\\", "/");   // IMPORTANT FIX
        String json = "{ \"folder\": \"" + f + "\", \"file\": \"" + fileName + "\" }";

        System.out.println("[Producer] Sending -> " + json);
        producer.send(new ProducerRecord<>(topic, json));
    }

    public static void main(String[] args) throws Exception {
        new FileWatcherProducer().run();
    }
}
