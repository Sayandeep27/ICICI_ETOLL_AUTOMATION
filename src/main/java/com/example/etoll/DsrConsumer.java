package com.example.etoll;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.kafka.clients.consumer.ConsumerRecord;
import org.apache.kafka.clients.consumer.KafkaConsumer;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.util.Collections;
import java.util.Map;
import java.util.Properties;

public class DsrConsumer {

    private static final ObjectMapper mapper = new ObjectMapper();

    public static void main(String[] args) throws Exception {

        Properties props = KafkaConfig.getConsumerProps();
        KafkaConsumer<String, String> consumer = new KafkaConsumer<>(props);

        consumer.subscribe(Collections.singletonList("dsr_topic"));
        System.out.println("[Consumer] Started...");

        while (true) {
            for (ConsumerRecord<String, String> rec :
                    consumer.poll(Duration.ofSeconds(1))) {

                String json = rec.value();
                System.out.println("[Consumer] Received -> " + json);

                try {
                    // --------------------------------------------------
                    // SAFE JSON PARSING
                    // --------------------------------------------------
                    JsonNode node = mapper.readTree(json);
                    String folder = node.get("folder").asText();
                    String file = node.get("file").asText();

                    Path dsrPath = Paths.get(folder, file)
                            .toAbsolutePath()
                            .normalize();

                    System.out.println("[Consumer] Using file: " + dsrPath);

                    // --------------------------------------------------
                    // CALL GENERATOR
                    // --------------------------------------------------
                    Map<String, Object> result =
                            EtollVoucherGenerator.generateVoucher(dsrPath);

                    System.out.println("[Consumer] Generator result: " + result);
                }
                catch (Exception ex) {
                    System.out.println("[Consumer] ERROR processing message: "
                            + ex.getMessage());
                    ex.printStackTrace();
                }
            }
        }
    }
}
