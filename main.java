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
    public OfficeClawster.SheetLedger getSheetLedger() { return sheetLedger; }
    public OfficeClawster.InboxRegistry getInboxRegistry() { return inboxRegistry; }
    public OfficeClawster.ClawBotEngine getEngine() { return engine; }

    // -------------------------------------------------------------------------
    // OFFICE TASK TYPE
    // -------------------------------------------------------------------------

    public enum OfficeTaskType {
        WORD_DOC(0),
        EXCEL_SHEET(1),
        OUTLOOK_MAIL(2),
        POWERPOINT_SLIDE(3),
        ONENOTE_PAGE(4),
        ACCESS_DB(5),
        PUBLISHER_PUB(6),
        VISIO_DIAGRAM(7),
        PROJECT_PLAN(8),
        TEAMS_MSG(9),
        SHAREPOINT_ITEM(10),
        GENERIC_DOC(11),
        CALENDAR_EVENT(12),
        CONTACT_ENTRY(13),
        TASK_ITEM(14),
        UNKNOWN(15);

        private final int code;
        OfficeTaskType(int code) { this.code = code; }
        public int getCode() { return code; }
        public static OfficeTaskType fromCode(int c) {
            for (OfficeTaskType t : values()) if (t.code == c) return t;
            return UNKNOWN;
        }
    }

    // -------------------------------------------------------------------------
    // QUEUED DOCUMENT
    // -------------------------------------------------------------------------

    public static final class QueuedDocument {
        private final String docId;
        private final String enqueuedBy;
        private final OfficeTaskType docType;
        private final long queueEpoch;
        private final long enqueuedAtMs;
        private final String payloadHash;
