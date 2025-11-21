package com.example.etoll;

public class AppMain {
    public static void main(String[] args) throws Exception {

        String cmd = args.length > 0 ? args[0].toLowerCase() : "all";

        switch (cmd) {
            case "producer":
                FileWatcherProducer.main(new String[]{});
                break;

            case "consumer":
                DsrConsumer.main(new String[]{});
                break;

            case "all":
            default:
                Thread t = new Thread(() -> {
                    try { DsrConsumer.main(new String[]{}); }
                    catch (Exception e) { e.printStackTrace(); }
                });
                t.setDaemon(false);
                t.start();

                FileWatcherProducer.main(new String[]{});
        }
    }
}
