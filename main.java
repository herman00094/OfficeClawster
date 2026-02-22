/*
 * OfficeClawster — Clawbot engine for MS Office tasks. Single-file build; run main() for CLI or use programmatic API.
 * Domain: 0x1e4f7a9C2d5E8b0F3a6C9e1D4b7F0a3C6d9E2b5F8a1c4e7
 */

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.time.Instant;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

public final class OfficeClawster {

    public static final String CLAW_VERSION = "1.0.0-office";
    public static final int MAX_DOCS_PER_EPOCH = 256;
    public static final int CELL_SLOTS = 64;
    public static final int INBOX_SLOTS = 32;
    public static final int MAX_DOC_TYPE = 15;
    public static final int DEFAULT_QUEUE_CAP = 2048;

    private final OfficeClawster.DocQueue docQueue;
    private final OfficeClawster.SheetLedger sheetLedger;
    private final OfficeClawster.InboxRegistry inboxRegistry;
    private final OfficeClawster.ClawBotEngine engine;

    public OfficeClawster() {
        this.docQueue = new OfficeClawster.DocQueue(DEFAULT_QUEUE_CAP);
        this.sheetLedger = new OfficeClawster.SheetLedger(CELL_SLOTS);
        this.inboxRegistry = new OfficeClawster.InboxRegistry(INBOX_SLOTS);
        this.engine = new OfficeClawster.ClawBotEngine(docQueue, sheetLedger, inboxRegistry);
    }

    public OfficeClawster(int queueCap) {
        this.docQueue = new OfficeClawster.DocQueue(queueCap);
        this.sheetLedger = new OfficeClawster.SheetLedger(CELL_SLOTS);
        this.inboxRegistry = new OfficeClawster.InboxRegistry(INBOX_SLOTS);
        this.engine = new OfficeClawster.ClawBotEngine(docQueue, sheetLedger, inboxRegistry);
    }

    public OfficeClawster.DocQueue getDocQueue() { return docQueue; }
