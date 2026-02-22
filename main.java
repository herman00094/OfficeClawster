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
        private boolean processed;

        public QueuedDocument(String docId, String enqueuedBy, OfficeTaskType docType, long queueEpoch, String payloadHash) {
            this.docId = Objects.requireNonNull(docId);
            this.enqueuedBy = enqueuedBy != null ? enqueuedBy : "";
            this.docType = docType != null ? docType : OfficeTaskType.GENERIC_DOC;
            this.queueEpoch = queueEpoch;
            this.enqueuedAtMs = System.currentTimeMillis();
            this.payloadHash = payloadHash != null ? payloadHash : "";
            this.processed = false;
        }

        public String getDocId() { return docId; }
        public String getEnqueuedBy() { return enqueuedBy; }
        public OfficeTaskType getDocType() { return docType; }
        public long getQueueEpoch() { return queueEpoch; }
        public long getEnqueuedAtMs() { return enqueuedAtMs; }
        public String getPayloadHash() { return payloadHash; }
        public boolean isProcessed() { return processed; }
        public void setProcessed(boolean processed) { this.processed = processed; }
    }

    // -------------------------------------------------------------------------
    // SHEET CELL REF
    // -------------------------------------------------------------------------

    public static final class SheetCellRef {
        private final String cellRef;
        private final int sheetApp;
        private final long loggedAtMs;
        private final String valueHash;

        public SheetCellRef(String cellRef, int sheetApp, String valueHash) {
            this.cellRef = Objects.requireNonNull(cellRef);
            this.sheetApp = Math.max(0, Math.min(CELL_SLOTS - 1, sheetApp));
            this.loggedAtMs = System.currentTimeMillis();
            this.valueHash = valueHash != null ? valueHash : "";
        }

        public String getCellRef() { return cellRef; }
        public int getSheetApp() { return sheetApp; }
        public long getLoggedAtMs() { return loggedAtMs; }
        public String getValueHash() { return valueHash; }
    }

    // -------------------------------------------------------------------------
    // INBOX SLOT ITEM
    // -------------------------------------------------------------------------

    public static final class InboxSlotItem {
        private final String slotId;
        private final String reservedBy;
        private final int inboxType;
        private final long reservedAtMs;

        public InboxSlotItem(String slotId, String reservedBy, int inboxType) {
            this.slotId = Objects.requireNonNull(slotId);
            this.reservedBy = reservedBy != null ? reservedBy : "";
            this.inboxType = Math.max(0, Math.min(INBOX_SLOTS - 1, inboxType));
            this.reservedAtMs = System.currentTimeMillis();
        }

        public String getSlotId() { return slotId; }
        public String getReservedBy() { return reservedBy; }
        public int getInboxType() { return inboxType; }
        public long getReservedAtMs() { return reservedAtMs; }
    }

    // -------------------------------------------------------------------------
    // DOC QUEUE
    // -------------------------------------------------------------------------

    public static final class DocQueue {
        private final int capacity;
        private final Map<String, QueuedDocument> docs;
        private final List<String> docIdOrder;
        private long currentQueueEpoch;

        public DocQueue(int capacity) {
            this.capacity = Math.max(1, capacity);
            this.docs = new ConcurrentHashMap<>();
            this.docIdOrder = new ArrayList<>();
            this.currentQueueEpoch = 0;
        }

        public int getCapacity() { return capacity; }
        public int docCount() { return docs.size(); }
        public long getCurrentQueueEpoch() { return currentQueueEpoch; }
        public void bumpEpoch() { currentQueueEpoch++; }

        public Optional<QueuedDocument> getDoc(String docId) {
            return Optional.ofNullable(docs.get(docId));
        }

        public QueuedDocument enqueue(String docId, String enqueuedBy, OfficeTaskType docType, String payloadHash) {
            if (docs.size() >= capacity) throw new IllegalStateException("Queue full");
            if (docId == null || docId.isEmpty()) throw new IllegalArgumentException("Zero doc id");
            if (docs.containsKey(docId)) throw new IllegalStateException("Duplicate doc id");
            QueuedDocument d = new QueuedDocument(docId, enqueuedBy, docType, currentQueueEpoch, payloadHash);
            docs.put(docId, d);
            docIdOrder.add(docId);
            return d;
        }

        public void markProcessed(String docId) {
            QueuedDocument d = docs.get(docId);
            if (d == null) throw new IllegalArgumentException("Doc not found");
            if (d.isProcessed()) throw new IllegalStateException("Already processed");
            d.setProcessed(true);
        }

        public List<String> listDocIds() { return new ArrayList<>(docIdOrder); }
        public Collection<QueuedDocument> listDocs() { return new ArrayList<>(docs.values()); }
    }

    // -------------------------------------------------------------------------
    // SHEET LEDGER
    // -------------------------------------------------------------------------

    public static final class SheetLedger {
        private final int slots;
        private final Map<Integer, SheetCellRef> cellsBySlot;
        private final List<SheetCellRef> allCells;

        public SheetLedger(int slots) {
            this.slots = Math.max(1, slots);
            this.cellsBySlot = new ConcurrentHashMap<>();
            this.allCells = new ArrayList<>();
        }

        public void logCell(String cellRef, int sheetApp, String valueHash) {
            int slot = Math.abs(cellRef.hashCode()) % slots;
            while (cellsBySlot.containsKey(slot)) slot = (slot + 1) % slots;
            SheetCellRef ref = new SheetCellRef(cellRef, sheetApp, valueHash);
            cellsBySlot.put(slot, ref);
            allCells.add(ref);
        }

