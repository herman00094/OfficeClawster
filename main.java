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

        public Optional<SheetCellRef> getBySlot(int slot) {
            return Optional.ofNullable(cellsBySlot.get(slot));
        }

        public List<SheetCellRef> listCells() { return new ArrayList<>(allCells); }
        public List<SheetCellRef> listBySheetApp(int sheetApp) {
            return allCells.stream().filter(c -> c.getSheetApp() == sheetApp).collect(Collectors.toList());
        }
    }

    // -------------------------------------------------------------------------
    // INBOX REGISTRY
    // -------------------------------------------------------------------------

    public static final class InboxRegistry {
        private final int slots;
        private final Map<String, InboxSlotItem> bySlotId;
        private final Map<Integer, InboxSlotItem> bySlotIndex;
        private final List<InboxSlotItem> allSlots;

        public InboxRegistry(int slots) {
            this.slots = Math.max(1, slots);
            this.bySlotId = new ConcurrentHashMap<>();
            this.bySlotIndex = new ConcurrentHashMap<>();
            this.allSlots = new ArrayList<>();
        }

        public InboxSlotItem reserve(String slotId, String reservedBy, int inboxType) {
            if (bySlotId.containsKey(slotId)) throw new IllegalStateException("Slot id already reserved");
            if (bySlotIndex.size() >= slots) throw new IllegalStateException("Inbox slot cap");
            int idx = Math.abs(slotId.hashCode()) % slots;
            while (bySlotIndex.containsKey(idx)) idx = (idx + 1) % slots;
            InboxSlotItem item = new InboxSlotItem(slotId, reservedBy, inboxType);
            bySlotId.put(slotId, item);
            bySlotIndex.put(idx, item);
            allSlots.add(item);
            return item;
        }

        public Optional<InboxSlotItem> getBySlotId(String slotId) {
            return Optional.ofNullable(bySlotId.get(slotId));
        }

        public List<InboxSlotItem> listSlots() { return new ArrayList<>(allSlots); }
    }

    // -------------------------------------------------------------------------
    // CLAW BOT ENGINE
    // -------------------------------------------------------------------------

    public static final class ClawBotEngine {
        private final DocQueue docQueue;
        private final SheetLedger sheetLedger;
        private final InboxRegistry inboxRegistry;

        public ClawBotEngine(DocQueue docQueue, SheetLedger sheetLedger, InboxRegistry inboxRegistry) {
            this.docQueue = docQueue;
            this.sheetLedger = sheetLedger;
            this.inboxRegistry = inboxRegistry;
        }

        public List<QueuedDocument> getUnprocessedDocs() {
            return docQueue.listDocs().stream().filter(d -> !d.isProcessed()).collect(Collectors.toList());
        }

        public List<QueuedDocument> getDocsByType(OfficeTaskType type) {
            return docQueue.listDocs().stream().filter(d -> d.getDocType() == type).collect(Collectors.toList());
        }

        public List<QueuedDocument> getDocsInEpoch(long epoch) {
            return docQueue.listDocs().stream().filter(d -> d.getQueueEpoch() == epoch).collect(Collectors.toList());
        }

        public void processNextDoc() {
            Optional<QueuedDocument> next = docQueue.listDocs().stream().filter(d -> !d.isProcessed()).findFirst();
            next.ifPresent(d -> docQueue.markProcessed(d.getDocId()));
        }

        public String hashPayload(String payload) {
            return OfficeClawster.ExportUtils.hashContent(payload);
        }
    }

    // -------------------------------------------------------------------------
    // EXPORT / SERIALIZE
    // -------------------------------------------------------------------------

    public static final class ExportUtils {
        public static String toJson(OfficeClawster claw) {
            StringBuilder sb = new StringBuilder();
            sb.append("{\"version\":\"").append(CLAW_VERSION).append("\",\"epoch\":").append(claw.docQueue.getCurrentQueueEpoch()).append(",");
            sb.append("\"docs\":[");
            List<String> docIds = claw.docQueue.listDocIds();
            for (int i = 0; i < docIds.size(); i++) {
                QueuedDocument d = claw.docQueue.getDoc(docIds.get(i)).orElse(null);
                if (d == null) continue;
                if (i > 0) sb.append(",");
                sb.append("{\"id\":\"").append(escape(d.getDocId())).append("\",\"type\":\"").append(d.getDocType().name()).append("\",\"processed\":").append(d.isProcessed()).append("}");
            }
            sb.append("],\"cells\":").append(claw.sheetLedger.listCells().size()).append(",\"inboxSlots\":").append(claw.inboxRegistry.listSlots().size()).append("}");
            return sb.toString();
        }

        private static String escape(String s) {
            if (s == null) return "";
            return s.replace("\\", "\\\\").replace("\"", "\\\"");
        }

        public static String hashContent(String content) {
            if (content == null) return "";
            try {
                MessageDigest md = MessageDigest.getInstance("SHA-256");
                byte[] digest = md.digest(content.getBytes(StandardCharsets.UTF_8));
                StringBuilder hex = new StringBuilder();
                for (byte b : digest) hex.append(String.format("%02x", b));
                return hex.toString();
            } catch (NoSuchAlgorithmException e) {
                return Integer.toHexString(content.hashCode());
            }
        }

        public static void exportToFile(OfficeClawster claw, Path path) throws IOException {
            Files.write(path, toJson(claw).getBytes(StandardCharsets.UTF_8));
        }
    }

    // -------------------------------------------------------------------------
    // VALIDATION
    // -------------------------------------------------------------------------

    public static final class ValidationUtils {
        public static boolean isValidDocId(String docId) {
            return docId != null && !docId.isEmpty() && docId.length() <= 128;
        }

        public static boolean isValidCellRef(String cellRef) {
            return cellRef != null && cellRef.matches("[A-Za-z]+[0-9]+");
        }

        public static boolean isValidSlotId(String slotId) {
            return slotId != null && !slotId.isEmpty() && slotId.length() <= 64;
        }
    }

    // -------------------------------------------------------------------------
    // BATCH OPERATIONS
    // -------------------------------------------------------------------------

    public static final class BatchOps {
        public static int enqueueBatch(OfficeClawster claw, List<String> docIds, String enqueuedBy, OfficeTaskType type, String payloadHash) {
            int n = 0;
            for (String id : docIds) {
                if (!ValidationUtils.isValidDocId(id)) continue;
                try {
                    claw.docQueue.enqueue(id, enqueuedBy, type, payloadHash != null ? payloadHash : ExportUtils.hashContent(id));
                    n++;
                } catch (Exception ignored) { }
            }
            return n;
        }

        public static int logCellsBatch(OfficeClawster claw, List<String> cellRefs, int sheetApp) {
            int n = 0;
            for (String ref : cellRefs) {
                if (!ValidationUtils.isValidCellRef(ref)) continue;
                try {
                    claw.sheetLedger.logCell(ref, sheetApp, ExportUtils.hashContent(ref));
                    n++;
                } catch (Exception ignored) { }
            }
            return n;
        }
    }

    // -------------------------------------------------------------------------
    // ID GENERATOR
    // -------------------------------------------------------------------------

    public static final class IdGenerator {
        private static final String PREFIX_DOC = "doc-";
        private static final String PREFIX_CELL = "cell-";
        private static final String PREFIX_SLOT = "slot-";
        private int seq;

        public IdGenerator() { this.seq = 0; }
        public IdGenerator(int start) { this.seq = start; }

        public String nextDocId() { return PREFIX_DOC + Instant.now().toEpochMilli() + "-" + (seq++) + "-" + UUID.randomUUID().toString().substring(0, 6); }
        public String nextCellRef(String sheet, int row) { return PREFIX_CELL + sheet + row + "-" + (seq++); }
        public String nextSlotId() { return PREFIX_SLOT + (seq++) + "-" + Integer.toHexString((int) (System.nanoTime() & 0xFFFF)); }
    }

    // -------------------------------------------------------------------------
    // CLI MAIN
    // -------------------------------------------------------------------------

    public static void main(String[] args) throws IOException {
        OfficeClawster claw = new OfficeClawster(4096);
        IdGenerator idGen = new IdGenerator(100);

        if (args.length > 0) {
            switch (args[0].toLowerCase()) {
                case "export":
                    Path out = args.length > 1 ? Paths.get(args[1]) : Paths.get("office_clawster_export.json");
                    ExportUtils.exportToFile(claw, out);
                    System.out.println("Exported to " + out.toAbsolutePath());
                    return;
                case "stats":
                    System.out.println("Docs: " + claw.docQueue.docCount() + ", Cells: " + claw.sheetLedger.listCells().size() + ", Inbox: " + claw.inboxRegistry.listSlots().size());
                    System.out.println("Epoch: " + claw.docQueue.getCurrentQueueEpoch());
                    return;
                case "enqueue":
                    if (args.length >= 2) {
                        String docId = args[1];
                        OfficeTaskType type = args.length >= 3 ? OfficeTaskType.fromCode(Integer.parseInt(args[2])) : OfficeTaskType.WORD_DOC;
                        claw.docQueue.enqueue(docId, "cli", type, ExportUtils.hashContent(docId));
                        System.out.println("Enqueued " + docId);
                    }
                    return;
                default:
                    break;
            }
        }

        String d1 = idGen.nextDocId();
        String d2 = idGen.nextDocId();
        claw.docQueue.enqueue(d1, "system", OfficeTaskType.WORD_DOC, ExportUtils.hashContent("sample1"));
        claw.docQueue.enqueue(d2, "system", OfficeTaskType.EXCEL_SHEET, ExportUtils.hashContent("sample2"));
        claw.sheetLedger.logCell("A1", 0, ExportUtils.hashContent("value"));
        claw.sheetLedger.logCell("B2", 0, ExportUtils.hashContent("value2"));
        claw.inboxRegistry.reserve(idGen.nextSlotId(), "system", 0);

        System.out.println("OfficeClawster ready. Docs: " + claw.docQueue.docCount());
        System.out.println(ExportUtils.toJson(claw));
    }

    // -------------------------------------------------------------------------
    // COMMAND PARSER (CLI)
    // -------------------------------------------------------------------------

    public static final class CommandParser {
        private final OfficeClawster claw;
        private final IdGenerator idGen;

        public CommandParser(OfficeClawster claw) {
            this.claw = claw;
            this.idGen = new IdGenerator(500);
        }

        public boolean parse(String line) {
            if (line == null || line.trim().isEmpty()) return true;
            String[] parts = line.trim().split("\\s+");
            if (parts.length == 0) return true;
            switch (parts[0].toLowerCase()) {
                case "enqueue": return cmdEnqueue(parts);
                case "process": return cmdProcess(parts);
                case "logcell": return cmdLogCell(parts);
                case "reserve": return cmdReserve(parts);
                case "epoch": return cmdEpoch(parts);
                case "list": return cmdList(parts);
                case "export": return cmdExport(parts);
                case "quit": case "exit": return false;
                default: System.out.println("Unknown: " + parts[0]); return true;
            }
        }

        private boolean cmdEnqueue(String[] parts) {
            if (parts.length < 2) { System.out.println("Usage: enqueue <docId> [type]"); return true; }
            OfficeTaskType type = parts.length >= 3 ? OfficeTaskType.fromCode(Integer.parseInt(parts[2])) : OfficeTaskType.GENERIC_DOC;
            try {
                claw.docQueue.enqueue(parts[1], "user", type, ExportUtils.hashContent(parts[1]));
                System.out.println("Enqueued " + parts[1]);
            } catch (Exception e) { System.out.println("Error: " + e.getMessage()); }
            return true;
        }

        private boolean cmdProcess(String[] parts) {
            claw.getEngine().processNextDoc();
            System.out.println("Processed next doc");
            return true;
        }

        private boolean cmdLogCell(String[] parts) {
            if (parts.length < 2) { System.out.println("Usage: logcell <cellRef> [sheetApp]"); return true; }
            int app = parts.length >= 3 ? Integer.parseInt(parts[2]) : 0;
            try {
                claw.sheetLedger.logCell(parts[1], app, ExportUtils.hashContent(parts[1]));
                System.out.println("Logged " + parts[1]);
            } catch (Exception e) { System.out.println("Error: " + e.getMessage()); }
            return true;
        }

        private boolean cmdReserve(String[] parts) {
            if (parts.length < 2) { System.out.println("Usage: reserve <slotId> [inboxType]"); return true; }
            int type = parts.length >= 3 ? Integer.parseInt(parts[2]) : 0;
            try {
                claw.inboxRegistry.reserve(parts[1], "user", type);
                System.out.println("Reserved " + parts[1]);
            } catch (Exception e) { System.out.println("Error: " + e.getMessage()); }
            return true;
        }

        private boolean cmdEpoch(String[] parts) {
            claw.docQueue.bumpEpoch();
            System.out.println("Epoch: " + claw.docQueue.getCurrentQueueEpoch());
            return true;
        }

        private boolean cmdList(String[] parts) {
            if (parts.length < 2) { System.out.println("Usage: list docs|cells|slots"); return true; }
            switch (parts[1].toLowerCase()) {
                case "docs": claw.docQueue.listDocIds().forEach(id -> System.out.println("  " + id)); break;
                case "cells": claw.sheetLedger.listCells().forEach(c -> System.out.println("  " + c.getCellRef())); break;
                case "slots": claw.inboxRegistry.listSlots().forEach(s -> System.out.println("  " + s.getSlotId())); break;
                default: System.out.println("Unknown list type"); break;
            }
            return true;
        }

        private boolean cmdExport(String[] parts) {
            try {
                Path p = parts.length >= 2 ? Paths.get(parts[1]) : Paths.get("claw_export.json");
                ExportUtils.exportToFile(claw, p);
                System.out.println("Exported to " + p);
            } catch (Exception e) { System.out.println("Error: " + e.getMessage()); }
            return true;
        }
    }

    // -------------------------------------------------------------------------
    // REPORT BUILDER
    // -------------------------------------------------------------------------

    public static final class ReportBuilder {
        private final OfficeClawster claw;

        public ReportBuilder(OfficeClawster claw) { this.claw = claw; }

        public String buildSummary() {
            StringBuilder sb = new StringBuilder();
            sb.append("=== OfficeClawster Report ===\n");
            sb.append("Docs queued: ").append(claw.docQueue.docCount()).append("\n");
            sb.append("Unprocessed: ").append(claw.getEngine().getUnprocessedDocs().size()).append("\n");
            sb.append("Cells logged: ").append(claw.sheetLedger.listCells().size()).append("\n");
            sb.append("Inbox slots: ").append(claw.inboxRegistry.listSlots().size()).append("\n");
            sb.append("Current epoch: ").append(claw.docQueue.getCurrentQueueEpoch()).append("\n");
            return sb.toString();
        }

        public String buildDocTypeBreakdown() {
            StringBuilder sb = new StringBuilder();
            Map<OfficeTaskType, Long> counts = claw.docQueue.listDocs().stream()
                    .collect(Collectors.groupingBy(QueuedDocument::getDocType, Collectors.counting()));
            for (OfficeTaskType t : OfficeTaskType.values()) {
                if (counts.getOrDefault(t, 0L) > 0)
                    sb.append(t.name()).append(": ").append(counts.get(t)).append("\n");
            }
            return sb.toString();
        }
    }

    // -------------------------------------------------------------------------
    // EPOCH MANAGER
    // -------------------------------------------------------------------------

    public static final class EpochManager {
        private final DocQueue docQueue;
        public static final long MS_PER_EPOCH = 60_000;

        public EpochManager(DocQueue docQueue) { this.docQueue = docQueue; }

        public long getCurrentEpoch() { return docQueue.getCurrentQueueEpoch(); }
        public void advanceEpoch() { docQueue.bumpEpoch(); }
        public int docsInEpoch(long epoch) {
            return (int) docQueue.listDocs().stream().filter(d -> d.getQueueEpoch() == epoch).count();
        }
    }

    // -------------------------------------------------------------------------
    // OFFICE TASK EXECUTOR (STUB)
    // -------------------------------------------------------------------------

    public static final class OfficeTaskExecutor {
        private final ClawBotEngine engine;
        private int executedCount;

        public OfficeTaskExecutor(ClawBotEngine engine) {
            this.engine = engine;
            this.executedCount = 0;
        }

        public int getExecutedCount() { return executedCount; }
        public void executeNext() {
            List<QueuedDocument> unprocessed = engine.getUnprocessedDocs();
            if (!unprocessed.isEmpty()) {
                engine.docQueue.markProcessed(unprocessed.get(0).getDocId());
                executedCount++;
            }
        }
        public void executeAll() {
            while (!engine.getUnprocessedDocs().isEmpty()) executeNext();
        }
    }

    // -------------------------------------------------------------------------
    // FILTERS
    // -------------------------------------------------------------------------

    public static final class DocFilters {
        public static List<QueuedDocument> byType(Collection<QueuedDocument> docs, OfficeTaskType type) {
            return docs.stream().filter(d -> d.getDocType() == type).collect(Collectors.toList());
        }
        public static List<QueuedDocument> unprocessedOnly(Collection<QueuedDocument> docs) {
            return docs.stream().filter(d -> !d.isProcessed()).collect(Collectors.toList());
