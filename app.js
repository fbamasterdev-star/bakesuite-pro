/**
 * Bakery Business Manager Pro (BakeSuite Pro)
 * Senior Full-Stack Architecture - Single Page Application Core Engine
 * Complete Spreadsheet-Like Editing Engine, Reactive Math & High-DPI Visual Charts
 */

// ===================================================================
// DEFAULT TEMPLATE DATA (100% Replicated from Google Sheets CSVs)
// ===================================================================
const DEFAULT_BAKERY_DATA = {
  settings: {
    bakeryName: "BakeSuite Artisan Bakery",
    chefName: "Chef Elena Rostova",
    phone: "(555) 019-2834",
    email: "orders@bakesuitebakery.com",
    address: "142 Magnolia Lane, Suite B",
    currency: "$",
    laborRate: 25.00,
    overheadFactor: 15.0,
    targetMargin: 45.0,
    deliveryBaseFee: 15.00,
    mileageRate: 1.50,
    depositReqPercent: 50.0,
    salesTaxRate: 6.5,
    cfoLicense: "CFO Registration #4928-TX",
    baselineCogsPercent: 35.0,
    expenseShares: {
      ingredients: 38.0,
      labor: 42.0,
      overhead: 15.0,
      packaging: 5.0
    },
    bankingInfo: "Zelle: billing@bakesuitebakery.com | Venmo: @BakeSuite-Artisan\nDirect Bank Transfer: BakeSuite Artisan Bakery LLC | Routing: 122000496 | Acct: 982143098",
    stateDisclosure: "MADE IN A HOME KITCHEN THAT IS NOT SUBJECT TO STANDARD ROUTINE GOVERNMENT FOOD SAFETY INSPECTIONS.",
    safeHandling: "Keep refrigerated below 40°F (4°C). Bring to room temperature 1 hour before serving. Consume within 72 hours.",
    bakerSignature: "Made with love by Elena Rostova • BakeSuite Artisan Bakery"
  },
  unitConversions: [
    { id: 'u-1', unit: 'g', category: 'Mass', baseUnit: 'g', multiplier: 1.0 },
    { id: 'u-2', unit: 'kg', category: 'Mass', baseUnit: 'g', multiplier: 1000.0 },
    { id: 'u-3', unit: 'oz', category: 'Mass', baseUnit: 'g', multiplier: 28.3495 },
    { id: 'u-4', unit: 'lb', category: 'Mass', baseUnit: 'g', multiplier: 453.592 },
    { id: 'u-5', unit: 'ml', category: 'Volume', baseUnit: 'ml', multiplier: 1.0 },
    { id: 'u-6', unit: 'l', category: 'Volume', baseUnit: 'ml', multiplier: 1000.0 },
    { id: 'u-7', unit: 'cup', category: 'Volume', baseUnit: 'ml', multiplier: 240.0 },
    { id: 'u-8', unit: 'tbsp', category: 'Volume', baseUnit: 'ml', multiplier: 15.0 },
    { id: 'u-9', unit: 'tsp', category: 'Volume', baseUnit: 'ml', multiplier: 5.0 },
    { id: 'u-10', unit: 'unit', category: 'Count', baseUnit: 'unit', multiplier: 1.0 },
    { id: 'u-11', unit: 'dozen', category: 'Count', baseUnit: 'unit', multiplier: 12.0 }
  ],
  panMatrices: [
    { id: 'pan-1', dimension: '4" Round', shape: 'Round', party: 6, wedding: 8, area: 12.6, batter: 2.5, dowels: 0, board: '6" Cake Board' },
    { id: 'pan-2', dimension: '6" Round', shape: 'Round', party: 10, wedding: 14, area: 28.3, batter: 4.0, dowels: 0, board: '8" Cake Board' },
    { id: 'pan-3', dimension: '8" Round', shape: 'Round', party: 20, wedding: 26, area: 50.3, batter: 6.5, dowels: 3, board: '10" Cake Board' },
    { id: 'pan-4', dimension: '10" Round', shape: 'Round', party: 30, wedding: 38, area: 78.5, batter: 10.5, dowels: 4, board: '12" Cake Board' },
    { id: 'pan-5', dimension: '12" Round', shape: 'Round', party: 44, wedding: 56, area: 113.1, batter: 15.0, dowels: 6, board: '14" Cake Board' },
    { id: 'pan-6', dimension: '14" Round', shape: 'Round', party: 60, wedding: 78, area: 153.9, batter: 21.0, dowels: 8, board: '16" Cake Board' },
    { id: 'pan-7', dimension: '4" Square', shape: 'Square', party: 8, wedding: 10, area: 16.0, batter: 3.0, dowels: 0, board: '6" Square Board' },
    { id: 'pan-8', dimension: '6" Square', shape: 'Square', party: 12, wedding: 18, area: 36.0, batter: 5.0, dowels: 0, board: '8" Square Board' },
    { id: 'pan-9', dimension: '8" Square', shape: 'Square', party: 24, wedding: 32, area: 64.0, batter: 8.0, dowels: 4, board: '10" Square Board' },
    { id: 'pan-10', dimension: '10" Square', shape: 'Square', party: 38, wedding: 50, area: 100.0, batter: 13.0, dowels: 5, board: '12" Square Board' },
    { id: 'pan-11', dimension: '12" Square', shape: 'Square', party: 54, wedding: 72, area: 144.0, batter: 19.0, dowels: 7, board: '14" Square Board' },
    { id: 'pan-12', dimension: '14" Square', shape: 'Square', party: 74, wedding: 98, area: 196.0, batter: 26.0, dowels: 9, board: '16" Square Board' },
    { id: 'pan-13', dimension: '9x13" Quarter Sheet', shape: 'Rectangular', party: 24, wedding: 36, area: 117.0, batter: 14.0, dowels: 0, board: 'Quarter Sheet Board' },
    { id: 'pan-14', dimension: '12x18" Half Sheet', shape: 'Rectangular', party: 48, wedding: 72, area: 216.0, batter: 26.0, dowels: 0, board: 'Half Sheet Board' },
    { id: 'pan-15', dimension: 'Standard Cupcake (Dozen)', shape: 'Individual', party: 12, wedding: 12, area: 48.0, batter: 6.0, dowels: 0, board: 'Cupcake Box / Platter' }
  ],
  monthlyPerformance: [
    { id: 'm-1', month: 'Jul 2026', revenue: 2850.00, cogs: 997.50 },
    { id: 'm-2', month: 'Aug 2026', revenue: 3120.00, cogs: 1092.00 },
    { id: 'm-3', month: 'Sep 2026', revenue: 3450.00, cogs: 1207.50 },
    { id: 'm-4', month: 'Oct 2026', revenue: 3880.00, cogs: 1358.00 },
    { id: 'm-5', month: 'Nov 2026 (Est)', revenue: 4200.00, cogs: 1470.00 }
  ],
  inventory: [
    { id: 'inv-1', name: 'Flour (All Purpose)', category: 'Ingredient', packUnit: 'kg', packPrice: 12.50, packQty: 5, baseCost: 0.0025, stockQty: 8, reorderThreshold: 3 },
    { id: 'inv-2', name: 'Eggs', category: 'Ingredient', packUnit: 'dozen', packPrice: 19.50, packQty: 5, baseCost: 0.3250, stockQty: 6, reorderThreshold: 2 },
    { id: 'inv-3', name: 'Granulated Sugar', category: 'Ingredient', packUnit: 'kg', packPrice: 8.00, packQty: 4, baseCost: 0.0020, stockQty: 10, reorderThreshold: 4 },
    { id: 'inv-4', name: 'Unsalted Butter', category: 'Ingredient', packUnit: 'kg', packPrice: 16.00, packQty: 2, baseCost: 0.0080, stockQty: 5, reorderThreshold: 3 },
    { id: 'inv-5', name: 'Pure Vanilla Extract', category: 'Ingredient', packUnit: 'ml', packPrice: 25.00, packQty: 500, baseCost: 0.0500, stockQty: 3, reorderThreshold: 2 },
    { id: 'inv-6', name: 'Whole Milk', category: 'Ingredient', packUnit: 'l', packPrice: 4.50, packQty: 1, baseCost: 0.0045, stockQty: 12, reorderThreshold: 4 },
    { id: 'inv-7', name: 'Baking Powder', category: 'Ingredient', packUnit: 'g', packPrice: 6.00, packQty: 500, baseCost: 0.0120, stockQty: 4, reorderThreshold: 2 },
    { id: 'inv-8', name: 'Dutch Process Cocoa', category: 'Ingredient', packUnit: 'kg', packPrice: 14.00, packQty: 1, baseCost: 0.0140, stockQty: 3, reorderThreshold: 2 },
    { id: 'inv-9', name: 'Powdered Sugar', category: 'Ingredient', packUnit: 'kg', packPrice: 7.50, packQty: 3, baseCost: 0.0025, stockQty: 6, reorderThreshold: 3 },
    { id: 'inv-10', name: 'Heavy Whipping Cream', category: 'Ingredient', packUnit: 'l', packPrice: 6.20, packQty: 1, baseCost: 0.0062, stockQty: 8, reorderThreshold: 3 },
    { id: 'inv-11', name: 'Cream Cheese Block', category: 'Ingredient', packUnit: 'kg', packPrice: 11.50, packQty: 2, baseCost: 0.0077, stockQty: 4, reorderThreshold: 2 },
    { id: 'inv-12', name: 'Fine Sea Salt', category: 'Ingredient', packUnit: 'g', packPrice: 3.00, packQty: 1000, baseCost: 0.0030, stockQty: 5, reorderThreshold: 2 },
    { id: 'inv-13', name: 'Cake Box 8"', category: 'Packaging', packUnit: 'unit', packPrice: 24.00, packQty: 20, baseCost: 1.2000, stockQty: 15, reorderThreshold: 5 },
    { id: 'inv-14', name: 'Cake Board 8"', category: 'Packaging', packUnit: 'unit', packPrice: 15.00, packQty: 25, baseCost: 0.6000, stockQty: 20, reorderThreshold: 5 },
    { id: 'inv-15', name: 'Cake Box 10"', category: 'Packaging', packUnit: 'unit', packPrice: 28.00, packQty: 20, baseCost: 1.4000, stockQty: 8, reorderThreshold: 4 },
    { id: 'inv-16', name: 'Cake Board 10"', category: 'Packaging', packUnit: 'unit', packPrice: 18.00, packQty: 25, baseCost: 0.7200, stockQty: 12, reorderThreshold: 5 },
    { id: 'inv-17', name: 'Wooden Tier Dowels (12")', category: 'Sundries', packUnit: 'unit', packPrice: 9.00, packQty: 50, baseCost: 0.1800, stockQty: 35, reorderThreshold: 10 },
    { id: 'inv-18', name: 'Artisan Satin Ribbon (50yd)', category: 'Sundries', packUnit: 'unit', packPrice: 12.00, packQty: 1, baseCost: 12.0000, stockQty: 2, reorderThreshold: 2 }
  ],
  orders: [
    { id: 'ord-1', orderNumber: 'ORD-1001', customerName: 'Sarah Jenkins', phone: '(555) 234-5678', email: 'sarah.j@example.com', eventDate: '2026-10-20', eventTime: '14:00', items: 'Signature Vanilla Bean Cake (3-Tier)', status: 'Completed', deliveryFee: 15.00, totalQuoted: 220.00, depositPaid: 100.00, notes: 'Gold leaf accents, deliver to venue' },
    { id: 'ord-2', orderNumber: 'ORD-1002', customerName: 'Michael Chang', phone: '(555) 345-6789', email: 'mchang@example.com', eventDate: '2026-10-22', eventTime: '11:00', items: 'Chocolate Fudge Drip Cake (8")', status: 'Confirmed', deliveryFee: 0.00, totalQuoted: 85.00, depositPaid: 85.00, notes: "Pickup at bakery, 'Happy 30th Mike'" },
    { id: 'ord-3', orderNumber: 'ORD-1003', customerName: 'Emily Watson', phone: '(555) 456-7890', email: 'emily.w@example.com', eventDate: '2026-10-25', eventTime: '16:00', items: 'Berry Chantilly Naked Cake (10")', status: 'Completed', deliveryFee: 20.00, totalQuoted: 160.00, depositPaid: 80.00, notes: 'Fresh organic berries, gluten-free base' },
    { id: 'ord-4', orderNumber: 'ORD-1004', customerName: 'David Ross', phone: '(555) 567-8901', email: 'dross@example.com', eventDate: '2026-10-18', eventTime: '10:30', items: 'Assorted Cupcake Platter (36 pcs)', status: 'Ready for Pickup', deliveryFee: 0.00, totalQuoted: 95.00, depositPaid: 95.00, notes: '12 Vanilla, 12 Chocolate, 12 Red Velvet' },
    { id: 'ord-5', orderNumber: 'ORD-1005', customerName: 'Jessica Alba', phone: '(555) 678-9012', email: 'jalba@example.com', eventDate: '2026-10-28', eventTime: '15:00', items: 'Rustic Carrot Cake w/ Cream Cheese', status: 'Inquiry', deliveryFee: 10.00, totalQuoted: 75.00, depositPaid: 0.00, notes: 'Awaiting final flavor confirmation' },
    { id: 'ord-6', orderNumber: 'ORD-1006', customerName: 'Robert Miller', phone: '(555) 789-0123', email: 'rmiller@example.com', eventDate: '2026-10-14', eventTime: '12:00', items: 'Lemon Raspberry Celebration Cake', status: 'Completed', deliveryFee: 15.00, totalQuoted: 140.00, depositPaid: 140.00, notes: 'Paid in full upon pickup, 5-star review' },
    { id: 'ord-7', orderNumber: 'ORD-1007', customerName: 'Amanda Hayes', phone: '(555) 890-1234', email: 'ahayes@example.com', eventDate: '2026-10-30', eventTime: '13:30', items: 'Macaron Tower (50 pcs)', status: 'Confirmed', deliveryFee: 25.00, totalQuoted: 175.00, depositPaid: 50.00, notes: 'Pastel colors: pink, pistachio, lavender' },
    { id: 'ord-8', orderNumber: 'ORD-1008', customerName: 'Daniel Lee', phone: '(555) 901-2345', email: 'dlee@example.com', eventDate: '2026-10-15', eventTime: '17:00', items: 'Custom Birthday Sheet Cake', status: 'Cancelled', deliveryFee: 0.00, totalQuoted: 110.00, depositPaid: 0.00, notes: 'Client postponed event' },
    { id: 'ord-9', orderNumber: 'ORD-1009', customerName: 'Chloe Bennett', phone: '(555) 012-3456', email: 'cbennett@example.com', eventDate: '2026-11-05', eventTime: '14:00', items: 'Botanical Floral 2-Tier Cake (8"+6")', status: 'Baking', deliveryFee: 20.00, totalQuoted: 210.00, depositPaid: 105.00, notes: 'Pressed edible wildflowers & earl grey sponge' },
    { id: 'ord-10', orderNumber: 'ORD-1010', customerName: 'Marcus Vance', phone: '(555) 123-4567', email: 'mvance@example.com', eventDate: '2026-11-08', eventTime: '11:30', items: 'Salted Caramel Drip Cake (10")', status: 'Decorating', deliveryFee: 15.00, totalQuoted: 125.00, depositPaid: 62.50, notes: 'Dark chocolate drip with sea salt flakes' },
    { id: 'ord-11', orderNumber: 'ORD-1011', customerName: 'Olivia Taylor', phone: '(555) 234-9876', email: 'otaylor@example.com', eventDate: '2026-11-12', eventTime: '15:30', items: 'Gluten-Free Strawberry Shortcake (8")', status: 'Confirmed', deliveryFee: 0.00, totalQuoted: 90.00, depositPaid: 90.00, notes: 'Gluten-free certification requested' },
    { id: 'ord-12', orderNumber: 'ORD-1012', customerName: 'Lucas Gray', phone: '(555) 345-8765', email: 'lgray@example.com', eventDate: '2026-11-15', eventTime: '13:00', items: 'Gourmet Mini Tart Platter (48 pcs)', status: 'Inquiry', deliveryFee: 15.00, totalQuoted: 130.00, depositPaid: 0.00, notes: 'Corporate tasting inquiry' }
  ],
  recipes: [
    {
      id: 'rec-1',
      name: 'Signature Vanilla Bean Cake',
      category: 'Wedding Tier',
      yieldServings: 12,
      laborHours: 1.08333,
      laborRate: 25.00,
      overheadFactor: 15.0,
      targetMargin: 45.0,
      ingredients: [
        { name: 'Flour (All Purpose)', quantity: 350, unit: 'g', unitCost: 0.0025 },
        { name: 'Granulated Sugar', quantity: 300, unit: 'g', unitCost: 0.0020 },
        { name: 'Unsalted Butter', quantity: 225, unit: 'g', unitCost: 0.0080 },
        { name: 'Eggs', quantity: 4, unit: 'unit', unitCost: 0.3250 },
        { name: 'Whole Milk', quantity: 1, unit: 'cup', unitCost: 1.0800 },
        { name: 'Pure Vanilla Extract', quantity: 15, unit: 'ml', unitCost: 0.0500 },
        { name: 'Baking Powder', quantity: 12, unit: 'g', unitCost: 0.0120 },
        { name: 'Cake Board 8"', quantity: 1, unit: 'unit', unitCost: 0.6000 },
        { name: 'Cake Box 8"', quantity: 1, unit: 'unit', unitCost: 1.2000 }
      ]
    },
    {
      id: 'rec-2',
      name: 'Chocolate Fudge Drip Cake',
      category: 'Artisan Cake',
      yieldServings: 12,
      laborHours: 1.25,
      laborRate: 25.00,
      overheadFactor: 15.0,
      targetMargin: 45.0,
      ingredients: [
        { name: 'Flour (All Purpose)', quantity: 300, unit: 'g', unitCost: 0.0025 },
        { name: 'Granulated Sugar', quantity: 350, unit: 'g', unitCost: 0.0020 },
        { name: 'Dutch Process Cocoa', quantity: 100, unit: 'g', unitCost: 0.0140 },
        { name: 'Unsalted Butter', quantity: 200, unit: 'g', unitCost: 0.0080 },
        { name: 'Eggs', quantity: 4, unit: 'unit', unitCost: 0.3250 },
        { name: 'Whole Milk', quantity: 250, unit: 'ml', unitCost: 0.0045 },
        { name: 'Heavy Whipping Cream', quantity: 150, unit: 'ml', unitCost: 0.0062 },
        { name: 'Cake Board 8"', quantity: 1, unit: 'unit', unitCost: 0.6000 },
        { name: 'Cake Box 8"', quantity: 1, unit: 'unit', unitCost: 1.2000 }
      ]
    },
    {
      id: 'rec-3',
      name: 'Rustic Carrot Cake w/ Cream Cheese',
      category: 'Celebration',
      yieldServings: 14,
      laborHours: 1.40,
      laborRate: 25.00,
      overheadFactor: 15.0,
      targetMargin: 45.0,
      ingredients: [
        { name: 'Flour (All Purpose)', quantity: 350, unit: 'g', unitCost: 0.0025 },
        { name: 'Granulated Sugar', quantity: 250, unit: 'g', unitCost: 0.0020 },
        { name: 'Cream Cheese Block', quantity: 300, unit: 'g', unitCost: 0.0077 },
        { name: 'Powdered Sugar', quantity: 200, unit: 'g', unitCost: 0.0025 },
        { name: 'Unsalted Butter', quantity: 180, unit: 'g', unitCost: 0.0080 },
        { name: 'Eggs', quantity: 4, unit: 'unit', unitCost: 0.3250 },
        { name: 'Pure Vanilla Extract', quantity: 10, unit: 'ml', unitCost: 0.0500 },
        { name: 'Cake Board 10"', quantity: 1, unit: 'unit', unitCost: 0.7200 },
        { name: 'Cake Box 10"', quantity: 1, unit: 'unit', unitCost: 1.4000 }
      ]
    }
  ],
  tierPlanner: {
    targetGuests: 120,
    portionStandard: 'wedding',
    stackingMethod: 'Central + Perimeter Dowels',
    orderId: 'ord-1',
    tiers: [
      { id: 'tier-1', level: 'Tier 1 (Base Tier)', pan: '12" Round', height: 4.5, party: 44, wedding: 56, batter: 15.0, dowels: 6, board: '14" Cake Board' },
      { id: 'tier-2', level: 'Tier 2 (Middle Tier)', pan: '10" Round', height: 4.0, party: 30, wedding: 38, batter: 10.5, dowels: 4, board: '10" Cake Board' },
      { id: 'tier-3', level: 'Tier 3 (Top Tier)', pan: '6" Round', height: 4.0, party: 10, wedding: 14, batter: 4.0, dowels: 0, board: '6" Cake Board' },
      { id: 'tier-4', level: 'Tier 4 (Optional)', pan: '', height: 0, party: 0, wedding: 0, batter: 0, dowels: 0, board: '' },
      { id: 'tier-5', level: 'Tier 5 (Optional)', pan: '', height: 0, party: 0, wedding: 0, batter: 0, dowels: 0, board: '' }
    ]
  },
  groceryList: [
    { id: 'gr-1', name: 'Flour (All Purpose)', category: 'Ingredient', unit: 'kg', weeklyReq: 18, stockQty: 8, estCost: 25.00, bought: false },
    { id: 'gr-2', name: 'Granulated Sugar', category: 'Ingredient', unit: 'kg', weeklyReq: 14, stockQty: 10, estCost: 8.00, bought: false },
    { id: 'gr-3', name: 'Unsalted Butter', category: 'Ingredient', unit: 'kg', weeklyReq: 12, stockQty: 5, estCost: 56.00, bought: false },
    { id: 'gr-4', name: 'Eggs', category: 'Ingredient', unit: 'dozen', weeklyReq: 10, stockQty: 6, estCost: 15.60, bought: false },
    { id: 'gr-5', name: 'Pure Vanilla Extract', category: 'Ingredient', unit: 'ml', weeklyReq: 250, stockQty: 3, estCost: 12.35, bought: false },
    { id: 'gr-6', name: 'Whole Milk', category: 'Ingredient', unit: 'l', weeklyReq: 6, stockQty: 12, estCost: 0.00, bought: false },
    { id: 'gr-7', name: 'Dutch Process Cocoa', category: 'Ingredient', unit: 'kg', weeklyReq: 4, stockQty: 3, estCost: 14.00, bought: false },
    { id: 'gr-8', name: 'Powdered Sugar', category: 'Ingredient', unit: 'kg', weeklyReq: 8, stockQty: 6, estCost: 5.00, bought: false },
    { id: 'gr-9', name: 'Heavy Whipping Cream', category: 'Ingredient', unit: 'l', weeklyReq: 6, stockQty: 8, estCost: 0.00, bought: false },
    { id: 'gr-10', name: 'Cake Box 8"', category: 'Packaging', unit: 'unit', weeklyReq: 20, stockQty: 15, estCost: 6.00, bought: false },
    { id: 'gr-11', name: 'Cake Board 8"', category: 'Packaging', unit: 'unit', weeklyReq: 25, stockQty: 20, estCost: 3.00, bought: false },
    { id: 'gr-12', name: 'Wooden Tier Dowels (12")', category: 'Sundries', unit: 'unit', weeklyReq: 20, stockQty: 35, estCost: 0.00, bought: false }
  ],
  contractTerms: [
    { id: 'ct-1', title: '1. Retainer & Payment Terms:', body: 'A non-refundable retainer of 50% is required to secure the production calendar slot. The remaining balance must be paid in full no later than 48 hours prior to the scheduled delivery or collection date.' },
    { id: 'ct-2', title: '2. Cancellation & Postponements:', body: 'Written notice is required for all cancellations. Retainers are strictly non-refundable. Cancellations made within 14 days of the scheduled event forfeit 100% of the total order balance.' },
    { id: 'ct-3', title: '3. Delivery, Venue Setup & Temperature Liability:', body: 'BakeSuite is not responsible for structural damage occurring after delivery/handoff. Cakes contain real butter and must be maintained in a climate-controlled room below 70°F (21°C).' },
    { id: 'ct-4', title: '4. Allergen & Cross-Contamination Notice:', body: 'Products are crafted in a kitchen processing milk, eggs, wheat, tree nuts, and soy. The client acknowledges this advisory and is responsible for notifying guests with food allergies.' },
    { id: 'ct-5', title: '5. Structural Dowels & Floral Elements:', body: 'Tiered cakes contain food-grade wooden or plastic support dowels that must be removed prior to cutting. Fresh florals are decorative and must be removed before consumption.' }
  ],
  invoiceCustomItems: {},
  license: null
};

// ===================================================================
// APPLICATION ARCHITECTURE CORE
// ===================================================================
class BakeryApp {
  constructor() {
    this.data = null;
    this.license = null;
    this.fileHandle = null;
    this.deferredInstallPrompt = null;
    this.autoSaveTimer = null;
    this.isSaving = false;
    this.currentView = 'dashboard';
    this.activeRecipeId = 'rec-1';
    this.db = null;
    this.chartHoverState = {
      monthlyIndex: -1,
      monthlyX: 0,
      monthlyY: 0,
      donutSlice: -1,
      donutX: 0,
      donutY: 0
    };
  }

  async init() {
    console.log("Initializing BakeSuite Pro with Visual Graphs & Enhanced Table Architecture...");
    await this.initIndexedDB();
    await this.loadInitialData();
    await this.checkLicenseState();
    this.initPWA();
    this.initEventListeners();
    this.initChartInteractivity();
    this.initViewRouting();
    this.renderAllViews();
    this.checkSavedFileHandle();
    this.updateOnlineStatus();

    // Redraw charts on window resize
    window.addEventListener('resize', () => {
      if (this.currentView === 'dashboard') {
        this.renderDashboardCharts();
      }
    });
  }

  // -----------------------------------------------------------------
  // 1. IndexedDB Persistence
  // -----------------------------------------------------------------
  initIndexedDB() {
    return new Promise((resolve) => {
      const request = indexedDB.open('BakeSuiteMasterDB', 1);
      request.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains('appStore')) {
          db.createObjectStore('appStore');
        }
      };
      request.onsuccess = (e) => {
        this.db = e.target.result;
        resolve(this.db);
      };
      request.onerror = (e) => {
        console.error('IndexedDB error:', e);
        resolve(null);
      };
    });
  }

  idbSet(key, value) {
    if (!this.db) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const tx = this.db.transaction('appStore', 'readwrite');
      const store = tx.objectStore('appStore');
      store.put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  }

  idbGet(key) {
    if (!this.db) return Promise.resolve(null);
    return new Promise((resolve) => {
      const tx = this.db.transaction('appStore', 'readonly');
      const store = tx.objectStore('appStore');
      const req = store.get(key);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
    });
  }

  // -----------------------------------------------------------------
  // 2. State Loading
  // -----------------------------------------------------------------
  async loadInitialData() {
    try {
      let saved = await this.idbGet('bakeryMasterData');
      if (!saved) {
        const lsData = localStorage.getItem('bakeryMasterData');
        if (lsData) saved = JSON.parse(lsData);
      }

      if (saved && saved.orders && saved.inventory) {
        this.data = this.mergeWithDefaultSchema(saved);
      } else {
        this.data = JSON.parse(JSON.stringify(DEFAULT_BAKERY_DATA));
        await this.idbSet('bakeryMasterData', this.data);
        localStorage.setItem('bakeryMasterData', JSON.stringify(this.data));
      }

      this.activeRecipeId = this.data.recipes[0]?.id || 'rec-1';
    } catch (err) {
      console.error("State loading error:", err);
      this.data = JSON.parse(JSON.stringify(DEFAULT_BAKERY_DATA));
    }
  }

  mergeWithDefaultSchema(saved) {
    const d = JSON.parse(JSON.stringify(DEFAULT_BAKERY_DATA));
    return {
      ...d,
      ...saved,
      settings: { ...d.settings, ...(saved.settings || {}) },
      unitConversions: saved.unitConversions || d.unitConversions,
      panMatrices: saved.panMatrices || d.panMatrices,
      monthlyPerformance: saved.monthlyPerformance || d.monthlyPerformance,
      contractTerms: saved.contractTerms || d.contractTerms,
      invoiceCustomItems: saved.invoiceCustomItems || {},
      license: saved.license || d.license || null
    };
  }

  async checkSavedFileHandle() {
    try {
      const handle = await this.idbGet('masterFileHandle');
      if (handle) {
        this.fileHandle = handle;
        const name = handle.name || 'Bakery_Master_Data.json';
        this.updateSyncUI(`📁 Connected to ${name} (Ready)`, 'success');
        document.getElementById('cfgLinkedFileName').textContent = name;
      }
    } catch (err) {
      console.warn("Could not restore file handle:", err);
    }
  }

  // -----------------------------------------------------------------
  // 3. PWA Installation & Service Worker
  // -----------------------------------------------------------------
  initPWA() {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('./sw.js')
        .then(() => console.log('Offline Service Worker registered.'))
        .catch(err => console.warn('SW failed:', err));
    }

    const installBtn = document.getElementById('btnInstallApp');
    window.addEventListener('beforeinstallprompt', (e) => {
      e.preventDefault();
      this.deferredInstallPrompt = e;
      if (installBtn) installBtn.style.display = 'inline-flex';
    });

    if (installBtn) {
      installBtn.addEventListener('click', async () => {
        if (!this.deferredInstallPrompt) return;
        this.deferredInstallPrompt.prompt();
        const { outcome } = await this.deferredInstallPrompt.userChoice;
        console.log(`PWA install outcome: ${outcome}`);
        this.deferredInstallPrompt = null;
        installBtn.style.display = 'none';
      });
    }

    window.addEventListener('appinstalled', () => {
      if (installBtn) installBtn.style.display = 'none';
    });

    window.addEventListener('online', () => this.updateOnlineStatus());
    window.addEventListener('offline', () => this.updateOnlineStatus());
  }

  updateOnlineStatus() {
    const isOnline = navigator.onLine;
    const dot = document.getElementById('onlineDot');
    const txt = document.getElementById('onlineText');
    if (dot && txt) {
      dot.className = isOnline ? 'status-dot' : 'status-dot offline';
      txt.textContent = isOnline ? '100% Offline Ready' : 'Running Offline';
    }
  }

  // -----------------------------------------------------------------
  // 4. File System Access API & Single-File Overwrite Auto-Save
  // -----------------------------------------------------------------
  initEventListeners() {
    window.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        this.saveNow();
      }
    });

    document.getElementById('btnSaveNow')?.addEventListener('click', () => this.saveNow());
    document.getElementById('btnConnectFile')?.addEventListener('click', () => this.connectMasterJSON());
    document.getElementById('btnOpenFile')?.addEventListener('click', () => this.openExistingJSON());
    document.getElementById('btnDownloadBackup')?.addEventListener('click', () => this.downloadMasterJSON());

    // Search and filter listeners
    document.getElementById('orderSearchInput')?.addEventListener('input', () => this.renderOrders());
    document.getElementById('orderStatusFilter')?.addEventListener('change', () => this.renderOrders());
    document.getElementById('orderPaymentFilter')?.addEventListener('change', () => this.renderOrders());

    document.getElementById('inventorySearchInput')?.addEventListener('input', () => this.renderInventory());
    document.getElementById('inventoryCategoryFilter')?.addEventListener('change', () => this.renderInventory());
    document.getElementById('inventoryLowStockOnly')?.addEventListener('change', () => this.renderInventory());

    document.getElementById('recipeSelectDropdown')?.addEventListener('change', (e) => {
      this.activeRecipeId = e.target.value;
      this.renderActiveRecipe();
    });

    document.getElementById('batchMultiplierInput')?.addEventListener('input', () => {
      this.calculateScaler();
    });

    document.getElementById('recipeYieldInput')?.addEventListener('input', () => this.handleRecipeParamChange());
    document.getElementById('recipeLaborHoursInput')?.addEventListener('input', () => this.handleRecipeParamChange());
    document.getElementById('recipeLaborRateInput')?.addEventListener('input', () => this.handleRecipeParamChange());
    document.getElementById('recipeOverheadInput')?.addEventListener('input', () => this.handleRecipeParamChange());
    document.getElementById('recipeMarginInput')?.addEventListener('input', () => this.handleRecipeParamChange());

    // Tier planner inputs
    document.getElementById('tpTargetGuests')?.addEventListener('input', () => this.renderTierPlanner());
    document.getElementById('tpPortionStandard')?.addEventListener('change', () => this.renderTierPlanner());
    document.getElementById('tpStackingMethod')?.addEventListener('change', () => this.renderTierPlanner());
    document.getElementById('tpOrderSelect')?.addEventListener('change', (e) => {
      this.data.tierPlanner.orderId = e.target.value;
      this.triggerAutoSave();
    });

    document.getElementById('convOriginalPan')?.addEventListener('change', () => this.calculatePanConverter());
    document.getElementById('convTargetPan')?.addEventListener('change', () => this.calculatePanConverter());

    // Weekly prep date filters
    document.getElementById('prepWeekStart')?.addEventListener('change', () => this.renderWeeklyPrep());
    document.getElementById('prepWeekEnd')?.addEventListener('change', () => this.renderWeeklyPrep());

    // Document selects
    document.getElementById('invoiceOrderSelect')?.addEventListener('change', (e) => this.renderInvoice(e.target.value));
    document.getElementById('labelOrderSelect')?.addEventListener('change', (e) => this.renderCottageLabel(e.target.value));
    document.getElementById('contractOrderSelect')?.addEventListener('change', (e) => this.renderContract(e.target.value));
  }

  triggerAutoSave() {
    const isAutoSave = document.getElementById('toggleAutosave')?.checked ?? true;
    
    this.idbSet('bakeryMasterData', this.data);
    localStorage.setItem('bakeryMasterData', JSON.stringify(this.data));

    if (!isAutoSave) return;
    if (this.autoSaveTimer) clearTimeout(this.autoSaveTimer);

    this.autoSaveTimer = setTimeout(() => {
      this.saveNow(true);
    }, 1000);
  }

  async saveNow(isAuto = false) {
    if (this.isSaving) return;
    this.isSaving = true;

    const timeStr = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });

    await this.idbSet('bakeryMasterData', this.data);
    localStorage.setItem('bakeryMasterData', JSON.stringify(this.data));

    if (this.fileHandle) {
      try {
        this.updateSyncUI(`⏳ Overwriting ${this.fileHandle.name}...`, 'saving');
        
        const options = { mode: 'readwrite' };
        if ((await this.fileHandle.queryPermission(options)) !== 'granted') {
          if ((await this.fileHandle.requestPermission(options)) !== 'granted') {
            throw new Error('Permission denied');
          }
        }

        const writable = await this.fileHandle.createWritable();
        const jsonString = JSON.stringify(this.data, null, 2);
        await writable.write(jsonString);
        await writable.close();

        this.updateSyncUI(`✅ Auto-saved to ${this.fileHandle.name} at ${timeStr}`, 'success');
        document.getElementById('cfgLinkedFileName').textContent = this.fileHandle.name;
      } catch (err) {
        console.warn("Direct file overwrite failed:", err);
        this.updateSyncUI(`⚠️ Saved locally at ${timeStr} (Click Connect Master JSON)`, 'unsaved');
      }
    } else {
      this.updateSyncUI(`💾 Saved locally at ${timeStr} (IndexedDB active)`, 'success');
    }

    this.isSaving = false;
  }

  async connectMasterJSON() {
    if (!('showSaveFilePicker' in window)) {
      alert("Your browser does not support the File System Access API. Use Download/Upload JSON.");
      return;
    }

    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: 'Bakery_Master_Data.json',
        types: [{ description: 'JSON Document', accept: { 'application/json': ['.json'] } }]
      });

      this.fileHandle = handle;
      await this.idbSet('masterFileHandle', handle);
      await this.saveNow();
      alert(`Master JSON linked: ${handle.name}\nAll future changes will automatically overwrite this single file!`);
    } catch (err) {
      if (err.name !== 'AbortError') alert("Failed to connect file: " + err.message);
    }
  }

  async openExistingJSON() {
    if (!('showOpenFilePicker' in window)) {
      document.getElementById('fileUploadInput')?.click();
      return;
    }

    try {
      const [handle] = await window.showOpenFilePicker({
        types: [{ description: 'JSON Document', accept: { 'application/json': ['.json'] } }],
        multiple: false
      });

      const file = await handle.getFile();
      const text = await file.text();
      const parsed = JSON.parse(text);

      if (!parsed.orders || !parsed.inventory) {
        alert("Selected JSON does not match the BakeSuite master schema.");
        return;
      }

      this.data = this.mergeWithDefaultSchema(parsed);
      this.fileHandle = handle;
      await this.idbSet('masterFileHandle', handle);
      await this.idbSet('bakeryMasterData', this.data);
      localStorage.setItem('bakeryMasterData', JSON.stringify(this.data));

      if (parsed.license && parsed.license.activated) {
        this.saveLicense(parsed.license);
      }

      this.renderAllViews();
      this.updateSyncUI(`✅ Loaded & Linked to ${handle.name}`, 'success');
      document.getElementById('cfgLinkedFileName').textContent = handle.name;
      alert(`Loaded successfully from ${handle.name}!`);
    } catch (err) {
      if (err.name !== 'AbortError') alert("Failed to open file: " + err.message);
    }
  }

  downloadMasterJSON() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(this.data, null, 2));
    const dlAnchor = document.createElement('a');
    dlAnchor.setAttribute("href", dataStr);
    dlAnchor.setAttribute("download", `Bakery_Master_Data_${new Date().toISOString().slice(0, 10)}.json`);
    document.body.appendChild(dlAnchor);
    dlAnchor.click();
    dlAnchor.remove();
  }

  importJSONFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const parsed = JSON.parse(e.target.result);
        if (!parsed.orders || !parsed.inventory) {
          alert("Invalid schema format.");
          return;
        }
        this.data = this.mergeWithDefaultSchema(parsed);
        this.idbSet('bakeryMasterData', this.data);
        localStorage.setItem('bakeryMasterData', JSON.stringify(this.data));
        if (parsed.license && parsed.license.activated) {
          this.saveLicense(parsed.license);
        }
        this.renderAllViews();
        alert("Data successfully restored from backup!");
      } catch (err) {
        alert("Error parsing backup: " + err.message);
      }
    };
    reader.readAsText(file);
    event.target.value = '';
  }

  resetToTemplate() {
    if (confirm("Reset all bakery data back to default Google Sheet template?")) {
      this.data = JSON.parse(JSON.stringify(DEFAULT_BAKERY_DATA));
      this.idbSet('bakeryMasterData', this.data);
      localStorage.setItem('bakeryMasterData', JSON.stringify(this.data));
      this.renderAllViews();
      this.saveNow();
      alert("Reset completed.");
    }
  }

  updateSyncUI(text, status) {
    const txt = document.getElementById('syncStatusText');
    const pulse = document.getElementById('syncPulse');
    if (txt) txt.textContent = text;
    if (pulse) {
      pulse.className = 'sync-pulse';
      if (status === 'saving') pulse.classList.add('saving');
      if (status === 'unsaved') pulse.classList.add('unsaved');
    }
  }

  // -----------------------------------------------------------------
  // 5. Navigation & Routing
  // -----------------------------------------------------------------
  initViewRouting() {
    document.querySelectorAll('.nav-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const viewId = btn.getAttribute('data-view');
        this.switchView(viewId);
      });
    });
  }

  switchView(viewId) {
    this.currentView = viewId;
    document.querySelectorAll('.nav-btn').forEach(b => {
      b.classList.toggle('active', b.getAttribute('data-view') === viewId);
    });

    document.querySelectorAll('.view-panel').forEach(panel => {
      panel.classList.toggle('active', panel.id === `view-${viewId}`);
    });

    if (viewId === 'dashboard') {
      this.renderDashboard();
      setTimeout(() => this.renderDashboardCharts(), 50);
    }
    if (viewId === 'orders') this.renderOrders();
    if (viewId === 'recipes') this.renderActiveRecipe();
    if (viewId === 'inventory') this.renderInventory();
    if (viewId === 'tierplanner') this.renderTierPlanner();
    if (viewId === 'weeklyprep') this.renderWeeklyPrep();
    if (viewId === 'invoice') this.populateInvoiceDropdown();
    if (viewId === 'cottagelabels') this.populateLabelDropdown();
    if (viewId === 'contracts') this.populateContractDropdown();
    if (viewId === 'settings') this.renderSettings();
  }

  renderAllViews() {
    this.renderDashboard();
    this.renderDashboardCharts();
    this.renderOrders();
    this.populateRecipeDropdown();
    this.renderActiveRecipe();
    this.renderInventory();
    this.renderTierPlanner();
    this.renderWeeklyPrep();
    this.populateInvoiceDropdown();
    this.populateLabelDropdown();
    this.populateContractDropdown();
    this.renderSettings();
  }

  // =================================================================
  // MODULE 1: EXECUTIVE DASHBOARD & VISUAL CHARTS
  // =================================================================
  renderDashboard() {
    const orders = this.data.orders;
    
    // Total Gross Booked Sales ($1,615.00)
    const grossRev = orders.reduce((sum, o) => sum + (parseFloat(o.totalQuoted) || 0), 0);
    
    // Baseline COGS % (Default 35.0%)
    const baseCogsPct = parseFloat(this.data.settings.baselineCogsPercent) || 35.0;
    const cogs = grossRev * (baseCogsPct / 100);
    const netProfit = grossRev - cogs;
    const netMargin = grossRev > 0 ? ((netProfit / grossRev) * 100) : 0;
    
    // Unpaid Receivables ($807.50)
    const unpaid = orders.reduce((sum, o) => {
      const bal = (parseFloat(o.totalQuoted) || 0) - (parseFloat(o.depositPaid) || 0);
      return sum + (bal > 0 ? bal : 0);
    }, 0);

    // Active Orders (all except Completed and Cancelled = 8)
    const activeOrders = orders.filter(o => !['Completed', 'Cancelled'].includes(o.status));

    document.getElementById('kpiGrossRevenue').textContent = `$${grossRev.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById('kpiCOGS').textContent = `$${cogs.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById('kpiCOGSPercent').textContent = `${baseCogsPct.toFixed(1)}% baseline food & materials`;
    document.getElementById('cfgBaselineCogsInput').value = baseCogsPct;
    document.getElementById('kpiNetProfit').textContent = `$${netProfit.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById('kpiNetMargin').textContent = `${netMargin.toFixed(1)}% gross margin`;
    document.getElementById('kpiUnpaidReceivables').textContent = `$${unpaid.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById('kpiActiveOrders').textContent = activeOrders.length;

    // Upcoming events table
    const upcoming = [...orders]
      .filter(o => o.status !== 'Cancelled')
      .sort((a, b) => new Date(a.eventDate) - new Date(b.eventDate))
      .slice(0, 6);

    const upBody = document.getElementById('dashboardUpcomingTableBody');
    if (upBody) {
      upBody.innerHTML = upcoming.map(o => {
        const bal = (parseFloat(o.totalQuoted) || 0) - (parseFloat(o.depositPaid) || 0);
        return `
          <tr>
            <td><strong style="color:var(--accent-gold); font-family:var(--font-mono);">${o.orderNumber}</strong></td>
            <td><strong>${o.customerName}</strong></td>
            <td>${o.eventDate}</td>
            <td>${o.eventTime || '12:00'}</td>
            <td>${o.items}</td>
            <td><span class="badge ${this.getStatusBadgeClass(o.status)}">${o.status}</span></td>
            <td class="text-right" style="font-weight:700; color:${bal > 0 ? 'var(--color-danger)' : 'var(--color-success)'}">$${bal.toFixed(2)}</td>
          </tr>
        `;
      }).join('');
    }

    // Pipeline Breakdown
    const stages = ['Inquiry', 'Confirmed', 'Baking', 'Decorating', 'Ready for Pickup', 'Completed', 'Cancelled'];
    const pCounts = {};
    stages.forEach(s => pCounts[s] = 0);
    orders.forEach(o => { if (pCounts[o.status] !== undefined) pCounts[o.status]++; });

    const pipelineContainer = document.getElementById('dashboardPipelineCounts');
    if (pipelineContainer) {
      pipelineContainer.innerHTML = stages.map(st => `
        <div style="display:flex; justify-content:space-between; align-items:center; padding:6px 10px; background:var(--bg-main); border-radius:var(--radius-sm);">
          <span style="font-size:0.84rem;">${st}</span>
          <span class="badge ${this.getStatusBadgeClass(st)}">${pCounts[st]}</span>
        </div>
      `).join('') + `
        <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 10px; border-top:1px solid var(--border-subtle); margin-top:4px;">
          <strong style="font-size:0.86rem;">Total Tracked Orders:</strong>
          <strong style="color:var(--accent-gold); font-size:1rem;">${orders.length}</strong>
        </div>
      `;
    }

    // Editable Expense Allocation Shares
    const sh = this.data.settings.expenseShares;
    const ingrAmt = cogs * (sh.ingredients / 100);
    const laborAmt = cogs * (sh.labor / 100);
    const overAmt = cogs * (sh.overhead / 100);
    const packAmt = cogs * (sh.packaging / 100);
    const totShare = sh.ingredients + sh.labor + sh.overhead + sh.packaging;

    document.getElementById('expIngredientsText').textContent = `${sh.ingredients.toFixed(1)}% ($${ingrAmt.toFixed(2)})`;
    document.getElementById('expIngredientsBar').style.width = `${sh.ingredients}%`;
    document.getElementById('expLaborText').textContent = `${sh.labor.toFixed(1)}% ($${laborAmt.toFixed(2)})`;
    document.getElementById('expLaborBar').style.width = `${sh.labor}%`;
    document.getElementById('expOverheadText').textContent = `${sh.overhead.toFixed(1)}% ($${overAmt.toFixed(2)})`;
    document.getElementById('expOverheadBar').style.width = `${sh.overhead}%`;
    document.getElementById('expPackagingText').textContent = `${sh.packaging.toFixed(1)}% ($${packAmt.toFixed(2)})`;
    document.getElementById('expPackagingBar').style.width = `${sh.packaging}%`;
    document.getElementById('expTotalText').textContent = `${totShare.toFixed(1)}% ($${cogs.toFixed(2)})`;

    // Top 5 Profitable Recipes
    const topBody = document.getElementById('dashboardTopRecipesBody');
    if (topBody) {
      const topRecipes = [
        { rank: '#1', name: 'Signature Vanilla Bean Cake (3-Tier)', cat: 'Wedding Tier', price: '$220.00', cogs: '$73.37', margin: '66.6%' },
        { rank: '#2', name: 'Botanical Floral 2-Tier Cake', cat: 'Celebration', price: '$210.00', cogs: '$68.50', margin: '67.4%' },
        { rank: '#3', name: 'Berry Chantilly Naked Cake (10")', cat: 'Artisan Cake', price: '$160.00', cogs: '$48.20', margin: '69.9%' },
        { rank: '#4', name: 'Salted Caramel Drip Cake (10")', cat: 'Artisan Cake', price: '$125.00', cogs: '$37.50', margin: '70.0%' },
        { rank: '#5', name: 'Assorted Cupcake Platter (36 pcs)', cat: 'Dessert Table', price: '$95.00', cogs: '$26.80', margin: '71.8%' }
      ];
      topBody.innerHTML = topRecipes.map(r => `
        <tr>
          <td><strong style="color:var(--accent-gold);">${r.rank}</strong></td>
          <td>${r.name}</td>
          <td><span style="color:var(--text-muted);">${r.cat}</span></td>
          <td class="text-right">${r.price}</td>
          <td class="text-right">${r.cogs}</td>
          <td class="text-right"><strong style="color:var(--color-success);">${r.margin}</strong></td>
        </tr>
      `).join('');
    }

    // Editable Monthly Performance Summary
    const monthlyBody = document.getElementById('dashboardMonthlyBody');
    if (monthlyBody) {
      monthlyBody.innerHTML = this.data.monthlyPerformance.map((m, idx) => {
        const net = (parseFloat(m.revenue) || 0) - (parseFloat(m.cogs) || 0);
        return `
          <tr>
            <td>
              <input type="text" class="inline-input" value="${m.month}" onchange="App.updateMonthlyPerformance(${idx}, 'month', this.value)">
            </td>
            <td class="text-right">
              <input type="number" step="10" class="inline-input text-right" value="${m.revenue}" onchange="App.updateMonthlyPerformance(${idx}, 'revenue', this.value)">
            </td>
            <td class="text-right">
              <input type="number" step="10" class="inline-input text-right" value="${m.cogs}" onchange="App.updateMonthlyPerformance(${idx}, 'cogs', this.value)">
            </td>
            <td class="text-right" style="color:var(--color-success); font-weight:700;">$${net.toFixed(2)}</td>
            <td class="text-center">
              <button class="btn btn-outline btn-sm" onclick="App.deleteMonthlyRow(${idx})">🗑️</button>
            </td>
          </tr>
        `;
      }).join('');
    }

    document.getElementById('navOrderCount').textContent = activeOrders.length;
    const lowStock = this.data.inventory.filter(i => i.stockQty <= i.reorderThreshold).length;
    document.getElementById('navLowStockCount').textContent = lowStock;
  }

  // =================================================================
  // CANVAS CHARTING ENGINE (Crisp High-DPI, Interactive & Self-Contained)
  // =================================================================
  initChartInteractivity() {
    const monthlyCanvas = document.getElementById('chartMonthlyTrends');
    if (monthlyCanvas) {
      const handleMonthlyMove = (e) => {
        const rect = monthlyCanvas.getBoundingClientRect();
        const clientX = e.touches ? e.touches[0].clientX : e.clientX;
        const clientY = e.touches ? e.touches[0].clientY : e.clientY;
        const mx = clientX - rect.left;
        const my = clientY - rect.top;

        const padLeft = 55;
        const padRight = 20;
        const chartW = rect.width - padLeft - padRight;
        const n = this.data.monthlyPerformance.length;
        if (n === 0) return;
        const groupW = chartW / n;

        let hoveredIdx = -1;
        if (mx >= padLeft && mx <= rect.width - padRight) {
          hoveredIdx = Math.floor((mx - padLeft) / groupW);
          if (hoveredIdx < 0 || hoveredIdx >= n) hoveredIdx = -1;
        }

        if (hoveredIdx !== this.chartHoverState.monthlyIndex) {
          this.chartHoverState.monthlyIndex = hoveredIdx;
          this.chartHoverState.monthlyX = mx;
          this.chartHoverState.monthlyY = my;
          this.drawMonthlyTrendsChart();
        }
      };

      const handleMonthlyLeave = () => {
        if (this.chartHoverState.monthlyIndex !== -1) {
          this.chartHoverState.monthlyIndex = -1;
          this.drawMonthlyTrendsChart();
        }
      };

      monthlyCanvas.addEventListener('mousemove', handleMonthlyMove);
      monthlyCanvas.addEventListener('mouseleave', handleMonthlyLeave);
      monthlyCanvas.addEventListener('touchmove', handleMonthlyMove, { passive: true });
      monthlyCanvas.addEventListener('touchend', handleMonthlyLeave);
    }

    const donutCanvas = document.getElementById('chartExpenseDonut');
    if (donutCanvas) {
      const handleDonutMove = (e) => {
        const rect = donutCanvas.getBoundingClientRect();
        const clientX = e.touches ? e.touches[0].clientX : e.clientX;
        const clientY = e.touches ? e.touches[0].clientY : e.clientY;
        const mx = clientX - rect.left;
        const my = clientY - rect.top;

        const centerX = rect.width / 2;
        const centerY = rect.height / 2;
        const radius = Math.min(centerX, centerY) - 15;
        const innerRadius = radius * 0.65;

        const dx = mx - centerX;
        const dy = my - centerY;
        const dist = Math.sqrt(dx * dx + dy * dy);

        let hoveredSlice = -1;
        if (dist >= innerRadius * 0.75 && dist <= radius + 20) {
          let angle = Math.atan2(dy, dx);
          let normAngle = angle - (-Math.PI / 2);
          if (normAngle < 0) normAngle += Math.PI * 2;

          const sh = this.data.settings.expenseShares;
          const slices = [
            { key: 'ingredients', share: sh.ingredients },
            { key: 'labor', share: sh.labor },
            { key: 'overhead', share: sh.overhead },
            { key: 'packaging', share: sh.packaging }
          ];
          const total = slices.reduce((acc, s) => acc + s.share, 0);

          let currentAngle = 0;
          for (let i = 0; i < slices.length; i++) {
            const sliceAngle = (slices[i].share / total) * (Math.PI * 2);
            if (normAngle >= currentAngle && normAngle <= currentAngle + sliceAngle) {
              hoveredSlice = i;
              break;
            }
            currentAngle += sliceAngle;
          }
        }

        if (hoveredSlice !== this.chartHoverState.donutSlice) {
          this.chartHoverState.donutSlice = hoveredSlice;
          this.chartHoverState.donutX = mx;
          this.chartHoverState.donutY = my;
          this.drawExpenseDonutChart();
        }
      };

      const handleDonutLeave = () => {
        if (this.chartHoverState.donutSlice !== -1) {
          this.chartHoverState.donutSlice = -1;
          this.drawExpenseDonutChart();
        }
      };

      donutCanvas.addEventListener('mousemove', handleDonutMove);
      donutCanvas.addEventListener('mouseleave', handleDonutLeave);
      donutCanvas.addEventListener('touchmove', handleDonutMove, { passive: true });
      donutCanvas.addEventListener('touchend', handleDonutLeave);
    }
  }

  renderDashboardCharts() {
    this.drawMonthlyTrendsChart();
    this.drawExpenseDonutChart();
  }

  drawMonthlyTrendsChart() {
    const canvas = document.getElementById('chartMonthlyTrends');
    if (!canvas) return;

    const ctx = canvas.getContext('2d');
    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();
    
    if (rect.width === 0 || rect.height === 0) return;

    canvas.width = rect.width * dpr;
    canvas.height = rect.height * dpr;
    ctx.scale(dpr, dpr);

    const width = rect.width;
    const height = rect.height;

    ctx.clearRect(0, 0, width, height);

    const data = this.data.monthlyPerformance;
    if (!data || data.length === 0) return;

    const padLeft = 55;
    const padRight = 20;
    const padTop = 25;
    const padBottom = 35;
    const chartW = width - padLeft - padRight;
    const chartH = height - padTop - padBottom;

    // Find max value
    let maxVal = 0;
    data.forEach(d => {
      const rev = parseFloat(d.revenue) || 0;
      if (rev > maxVal) maxVal = rev;
    });
    maxVal = Math.ceil((maxVal * 1.15) / 1000) * 1000 || 5000;

    // Draw horizontal grid lines
    ctx.strokeStyle = '#243250';
    ctx.lineWidth = 1;
    ctx.fillStyle = '#64748b';
    ctx.font = '10px Inter, sans-serif';
    ctx.textAlign = 'right';

    const gridLines = 4;
    for (let i = 0; i <= gridLines; i++) {
      const y = padTop + chartH - (i * (chartH / gridLines));
      const val = Math.round((i * (maxVal / gridLines)));

      ctx.beginPath();
      ctx.moveTo(padLeft, y);
      ctx.lineTo(width - padRight, y);
      ctx.stroke();

      ctx.fillText(`$${(val / 1000).toFixed(1)}k`, padLeft - 8, y + 3);
    }

    const n = data.length;
    const groupW = chartW / n;
    const barW = Math.min(24, (groupW - 14) / 2);

    const profitPoints = [];

    // Draw Bars for Sales and COGS
    data.forEach((d, i) => {
      const centerX = padLeft + (i * groupW) + (groupW / 2);
      const rev = parseFloat(d.revenue) || 0;
      const cogs = parseFloat(d.cogs) || 0;
      const net = rev - cogs;

      const revH = (rev / maxVal) * chartH;
      const cogsH = (cogs / maxVal) * chartH;

      const revX = centerX - barW - 2;
      const cogsX = centerX + 2;

      const revY = padTop + chartH - revH;
      const cogsY = padTop + chartH - cogsH;

      const isHovered = (this.chartHoverState.monthlyIndex === i);

      // Highlight background band if hovered
      if (isHovered) {
        ctx.fillStyle = 'rgba(245, 158, 11, 0.08)';
        ctx.fillRect(centerX - groupW / 2 + 3, padTop, groupW - 6, chartH);
        
        ctx.strokeStyle = 'rgba(245, 158, 11, 0.3)';
        ctx.lineWidth = 1;
        ctx.strokeRect(centerX - groupW / 2 + 3, padTop, groupW - 6, chartH);
      }

      // Revenue bar (Amber Gradient)
      const gradRev = ctx.createLinearGradient(0, revY, 0, padTop + chartH);
      gradRev.addColorStop(0, isHovered ? '#fbbf24' : '#f59e0b');
      gradRev.addColorStop(1, isHovered ? '#d97706' : '#b45309');
      ctx.fillStyle = gradRev;
      this.roundRect(ctx, revX, revY, barW, revH, 3);

      if (isHovered) {
        ctx.strokeStyle = '#fef3c7';
        ctx.lineWidth = 1.5;
        ctx.stroke();
      }

      // COGS bar (Slate Blue)
      const gradCogs = ctx.createLinearGradient(0, cogsY, 0, padTop + chartH);
      gradCogs.addColorStop(0, isHovered ? '#94a3b8' : '#64748b');
      gradCogs.addColorStop(1, isHovered ? '#475569' : '#334155');
      ctx.fillStyle = gradCogs;
      this.roundRect(ctx, cogsX, cogsY, barW, cogsH, 3);

      if (isHovered) {
        ctx.strokeStyle = '#e2e8f0';
        ctx.lineWidth = 1.5;
        ctx.stroke();
      }

      // Save net profit point for trendline
      const netY = padTop + chartH - ((net / maxVal) * chartH);
      profitPoints.push({ x: centerX, y: netY, val: net, isHovered });

      // X-axis label
      ctx.fillStyle = isHovered ? '#f59e0b' : '#94a3b8';
      ctx.font = isHovered ? 'bold 11px Inter, sans-serif' : '10px Inter, sans-serif';
      ctx.textAlign = 'center';
      ctx.fillText(d.month.replace(' 2026', ''), centerX, height - 12);
    });

    // Draw Profit Line
    if (profitPoints.length > 1) {
      ctx.strokeStyle = '#10b981';
      ctx.lineWidth = 3;
      ctx.shadowColor = 'rgba(16, 185, 129, 0.4)';
      ctx.shadowBlur = 8;

      ctx.beginPath();
      profitPoints.forEach((pt, i) => {
        if (i === 0) ctx.moveTo(pt.x, pt.y);
        else ctx.lineTo(pt.x, pt.y);
      });
      ctx.stroke();

      ctx.shadowBlur = 0;

      // Draw glowing circular points
      profitPoints.forEach(pt => {
        ctx.fillStyle = '#0f172a';
        ctx.beginPath();
        ctx.arc(pt.x, pt.y, pt.isHovered ? 7 : 5, 0, Math.PI * 2);
        ctx.fill();

        ctx.strokeStyle = pt.isHovered ? '#34d399' : '#10b981';
        ctx.lineWidth = pt.isHovered ? 3 : 2;
        ctx.beginPath();
        ctx.arc(pt.x, pt.y, pt.isHovered ? 7 : 5, 0, Math.PI * 2);
        ctx.stroke();

        // Data value pill above point
        ctx.fillStyle = pt.isHovered ? '#34d399' : '#10b981';
        ctx.font = pt.isHovered ? 'bold 11px Consolas, monospace' : 'bold 9px Consolas, monospace';
        ctx.textAlign = 'center';
        ctx.fillText(`+$${Math.round(pt.val)}`, pt.x, pt.y - (pt.isHovered ? 12 : 8));
      });
    }

    // DRAW INTERACTIVE HOVER TOOLTIP FLOATING CARD
    if (this.chartHoverState.monthlyIndex >= 0 && this.chartHoverState.monthlyIndex < n) {
      const d = data[this.chartHoverState.monthlyIndex];
      const rev = parseFloat(d.revenue) || 0;
      const cogs = parseFloat(d.cogs) || 0;
      const net = rev - cogs;
      const margin = rev > 0 ? ((net / rev) * 100).toFixed(1) : '0.0';

      const centerX = padLeft + (this.chartHoverState.monthlyIndex * groupW) + (groupW / 2);
      const tipW = 160;
      const tipH = 82;
      let tipX = centerX - tipW / 2;
      if (tipX < padLeft) tipX = padLeft;
      if (tipX + tipW > width - 10) tipX = width - tipW - 10;
      const tipY = padTop + 8;

      ctx.save();
      ctx.shadowColor = 'rgba(0, 0, 0, 0.7)';
      ctx.shadowBlur = 12;
      ctx.fillStyle = 'rgba(15, 23, 42, 0.96)';
      ctx.strokeStyle = 'rgba(245, 158, 11, 0.6)';
      ctx.lineWidth = 1;
      
      this.roundRect(ctx, tipX, tipY, tipW, tipH, 6);
      ctx.stroke();
      ctx.restore();

      ctx.textAlign = 'left';
      ctx.fillStyle = '#fff';
      ctx.font = 'bold 11px Inter, sans-serif';
      ctx.fillText(`📅 ${d.month}`, tipX + 10, tipY + 16);

      ctx.font = '10px Inter, sans-serif';
      ctx.fillStyle = '#fbbf24';
      ctx.fillText(`● Gross Sales: $${rev.toLocaleString('en-US', {minimumFractionDigits: 2})}`, tipX + 10, tipY + 34);

      ctx.fillStyle = '#94a3b8';
      ctx.fillText(`● Direct COGS: $${cogs.toLocaleString('en-US', {minimumFractionDigits: 2})}`, tipX + 10, tipY + 50);

      ctx.fillStyle = '#34d399';
      ctx.font = 'bold 10px Inter, sans-serif';
      ctx.fillText(`● Net Profit: +$${net.toLocaleString('en-US', {minimumFractionDigits: 2})} (${margin}%)`, tipX + 10, tipY + 67);
    }
  }

  drawExpenseDonutChart() {
    const canvas = document.getElementById('chartExpenseDonut');
    if (!canvas) return;

    const ctx = canvas.getContext('2d');
    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();

    if (rect.width === 0 || rect.height === 0) return;

    canvas.width = rect.width * dpr;
    canvas.height = rect.height * dpr;
    ctx.scale(dpr, dpr);

    const width = rect.width;
    const height = rect.height;

    ctx.clearRect(0, 0, width, height);

    const sh = this.data.settings.expenseShares;
    const slices = [
      { name: 'Direct Ingredients', short: 'Ingredients', share: sh.ingredients, color: '#f59e0b' },
      { name: 'Kitchen Labor', short: 'Labor', share: sh.labor, color: '#10b981' },
      { name: 'Overhead & Utilities', short: 'Overhead', share: sh.overhead, color: '#6366f1' },
      { name: 'Packaging Supplies', short: 'Supplies', share: sh.packaging, color: '#f43f5e' }
    ];

    const total = slices.reduce((acc, s) => acc + s.share, 0);
    const centerX = width / 2;
    const centerY = height / 2;
    const radius = Math.min(centerX, centerY) - 15;
    const innerRadius = radius * 0.65;

    const grossRev = this.data.orders.reduce((sum, o) => sum + (parseFloat(o.totalQuoted) || 0), 0);
    const cogs = grossRev * ((this.data.settings.baselineCogsPercent || 35) / 100);

    let startAngle = -Math.PI / 2;
    const hoveredSliceIdx = this.chartHoverState.donutSlice;

    slices.forEach((slice, idx) => {
      const sliceAngle = (slice.share / total) * (Math.PI * 2);
      const endAngle = startAngle + sliceAngle;
      const isHovered = (hoveredSliceIdx === idx);

      const rOuter = isHovered ? radius + 7 : radius;
      const rInner = isHovered ? innerRadius - 2 : innerRadius;

      ctx.save();
      if (isHovered) {
        ctx.shadowColor = slice.color;
        ctx.shadowBlur = 14;
      }

      ctx.beginPath();
      ctx.arc(centerX, centerY, rOuter, startAngle, endAngle);
      ctx.arc(centerX, centerY, rInner, endAngle, startAngle, true);
      ctx.closePath();

      ctx.fillStyle = slice.color;
      ctx.fill();

      // Divider gap
      ctx.strokeStyle = '#131b2e';
      ctx.lineWidth = isHovered ? 4 : 3;
      ctx.stroke();

      ctx.restore();

      startAngle = endAngle;
    });

    // Inner Circle cutout
    ctx.beginPath();
    ctx.arc(centerX, centerY, innerRadius - 3, 0, Math.PI * 2);
    ctx.fillStyle = '#0f172a';
    ctx.fill();

    // Center Dynamic Callout
    ctx.textAlign = 'center';

    if (hoveredSliceIdx >= 0 && hoveredSliceIdx < slices.length) {
      const hSlice = slices[hoveredSliceIdx];
      const hAmt = (hSlice.share / 100) * cogs;

      ctx.fillStyle = hSlice.color;
      ctx.font = 'bold 9px Inter, sans-serif';
      ctx.fillText(hSlice.short.toUpperCase(), centerX, centerY - 15);

      ctx.fillStyle = '#fff';
      ctx.font = 'bold 17px Consolas, monospace';
      ctx.fillText(`$${hAmt.toFixed(2)}`, centerX, centerY + 3);

      ctx.fillStyle = '#f59e0b';
      ctx.font = '10px Inter, sans-serif';
      ctx.fillText(`${hSlice.share.toFixed(1)}% Share`, centerX, centerY + 18);
    } else {
      ctx.fillStyle = '#64748b';
      ctx.font = '9px Inter, sans-serif';
      ctx.fillText('TOTAL EXPENSES', centerX, centerY - 14);

      ctx.fillStyle = '#fff';
      ctx.font = 'bold 16px Consolas, monospace';
      ctx.fillText(`$${cogs.toFixed(2)}`, centerX, centerY + 4);

      ctx.fillStyle = '#f59e0b';
      ctx.font = '10px Inter, sans-serif';
      ctx.fillText('100.0% Allocation', centerX, centerY + 18);
    }
  }

  roundRect(ctx, x, y, width, height, radius) {
    if (height < 2) height = 2;
    ctx.beginPath();
    ctx.moveTo(x + radius, y);
    ctx.lineTo(x + width - radius, y);
    ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
    ctx.lineTo(x + width, y + height);
    ctx.lineTo(x, y + height);
    ctx.lineTo(x, y + radius);
    ctx.quadraticCurveTo(x, y, x + radius, y);
    ctx.closePath();
    ctx.fill();
  }

  updateBaselineCogs(val) {
    this.data.settings.baselineCogsPercent = parseFloat(val) || 35.0;
    this.triggerAutoSave();
    this.renderDashboard();
    this.renderDashboardCharts();
  }

  updateExpenseShare(key, val) {
    if (this.data.settings.expenseShares[key] !== undefined) {
      this.data.settings.expenseShares[key] = parseFloat(val) || 0;
      this.triggerAutoSave();
      this.renderDashboard();
      this.renderDashboardCharts();
    }
  }

  updateMonthlyPerformance(index, field, value) {
    if (this.data.monthlyPerformance[index]) {
      this.data.monthlyPerformance[index][field] = field === 'month' ? value : (parseFloat(value) || 0);
      this.triggerAutoSave();
      this.renderDashboard();
      this.renderDashboardCharts();
    }
  }

  addMonthlyRow() {
    this.data.monthlyPerformance.push({
      id: 'm-' + Date.now(),
      month: 'New Month',
      revenue: 4000.00,
      cogs: 1400.00
    });
    this.triggerAutoSave();
    this.renderDashboard();
    this.renderDashboardCharts();
  }

  deleteMonthlyRow(index) {
    this.data.monthlyPerformance.splice(index, 1);
    this.triggerAutoSave();
    this.renderDashboard();
    this.renderDashboardCharts();
  }

  getStatusBadgeClass(status) {
    switch (status) {
      case 'Inquiry': return 'badge-inquiry';
      case 'Confirmed': return 'badge-confirmed';
      case 'Baking': return 'badge-baking';
      case 'Decorating': return 'badge-decorating';
      case 'Ready for Pickup': return 'badge-ready';
      case 'Completed': return 'badge-completed';
      case 'Cancelled': return 'badge-cancelled';
      default: return 'badge-inquiry';
    }
  }

  // =================================================================
  // MODULE 2: ORDERS & CRM (Wide, Unclipped Columns)
  // =================================================================
  renderOrders() {
    const search = document.getElementById('orderSearchInput')?.value.toLowerCase() || '';
    const statusFilter = document.getElementById('orderStatusFilter')?.value || 'ALL';
    const paymentFilter = document.getElementById('orderPaymentFilter')?.value || 'ALL';

    const tbody = document.getElementById('ordersTableBody');
    if (!tbody) return;

    let sumDelivery = 0;
    let sumTotal = 0;
    let sumDeposit = 0;
    let sumBalance = 0;

    const filtered = this.data.orders.filter(o => {
      const matchSearch = o.orderNumber.toLowerCase().includes(search) ||
                          o.customerName.toLowerCase().includes(search) ||
                          o.items.toLowerCase().includes(search);
      const matchStatus = statusFilter === 'ALL' || o.status === statusFilter;
      
      const bal = (parseFloat(o.totalQuoted) || 0) - (parseFloat(o.depositPaid) || 0);
      const payStatus = bal <= 0 ? 'Paid in Full' : (o.depositPaid > 0 ? 'Partial Deposit' : 'Unpaid');
      const matchPayment = paymentFilter === 'ALL' || payStatus === paymentFilter;

      return matchSearch && matchStatus && matchPayment;
    });

    this.data.orders.forEach(o => {
      const tot = parseFloat(o.totalQuoted) || 0;
      const dep = parseFloat(o.depositPaid) || 0;
      const del = parseFloat(o.deliveryFee) || 0;
      const bal = Math.max(0, tot - dep);
      sumDelivery += del;
      sumTotal += tot;
      sumDeposit += dep;
      sumBalance += bal;
    });

    document.getElementById('totDeliveryFee').textContent = `$${sumDelivery.toFixed(2)}`;
    document.getElementById('totTotalQuoted').textContent = `$${sumTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById('totDepositPaid').textContent = `$${sumDeposit.toFixed(2)}`;
    document.getElementById('totBalanceDue').textContent = `$${sumBalance.toFixed(2)}`;

    // Populate Separate Dedicated Total Pipeline Summary Segment
    const activeCount = this.data.orders.filter(o => !['Completed', 'Cancelled'].includes(o.status)).length;
    const rate = sumTotal > 0 ? ((sumDeposit / sumTotal) * 100).toFixed(1) : '0.0';

    const pipAct = document.getElementById('pipActiveCount');
    if (pipAct) pipAct.textContent = `${activeCount} Active Orders`;
    const pipRate = document.getElementById('pipRealizationRate');
    if (pipRate) pipRate.textContent = `${rate}% Realized`;

    const pipDel = document.getElementById('pipDeliveryFee');
    if (pipDel) pipDel.textContent = `$${sumDelivery.toFixed(2)}`;
    const pipTot = document.getElementById('pipTotalQuoted');
    if (pipTot) pipTot.textContent = `$${sumTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    const pipDep = document.getElementById('pipDepositPaid');
    if (pipDep) pipDep.textContent = `$${sumDeposit.toFixed(2)}`;
    const pipBal = document.getElementById('pipBalanceDue');
    if (pipBal) pipBal.textContent = `$${sumBalance.toFixed(2)}`;

    const pipProgLabel = document.getElementById('pipProgressLabel');
    if (pipProgLabel) pipProgLabel.textContent = `${rate}% ($${sumDeposit.toFixed(2)} / $${sumTotal.toFixed(2)})`;
    const pipProgBar = document.getElementById('pipProgressBar');
    if (pipProgBar) pipProgBar.style.width = `${Math.min(100, parseFloat(rate))}%`;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    tbody.innerHTML = filtered.map((o) => {
      const total = parseFloat(o.totalQuoted) || 0;
      const deposit = parseFloat(o.depositPaid) || 0;
      const balance = Math.max(0, total - deposit);
      const paymentStatus = balance <= 0 ? 'Paid in Full' : (deposit > 0 ? 'Partial Deposit' : 'Unpaid');
      const payBadgeClass = paymentStatus === 'Paid in Full' ? 'badge-paid' : (paymentStatus === 'Partial Deposit' ? 'badge-partial' : 'badge-unpaid');

      let daysLeft = '-';
      if (o.eventDate) {
        const evDate = new Date(o.eventDate);
        const diffDays = Math.ceil((evDate - today) / (1000 * 60 * 60 * 24));
        daysLeft = diffDays >= 0 ? `${diffDays}` : `${diffDays}`;
      }

      return `
        <tr>
          <td class="col-ord-num"><strong style="color:var(--accent-gold); font-family:var(--font-mono);">${o.orderNumber}</strong></td>
          <td class="col-ord-name">
            <input type="text" class="inline-input" value="${o.customerName}" title="${o.customerName}" onchange="App.updateOrderField('${o.id}', 'customerName', this.value)">
          </td>
          <td class="col-ord-phone">
            <input type="text" class="inline-input" style="font-size:0.8rem;" value="${o.phone || ''}" title="${o.phone || ''}" onchange="App.updateOrderField('${o.id}', 'phone', this.value)">
          </td>
          <td class="col-ord-email">
            <input type="text" class="inline-input" style="font-size:0.8rem;" value="${o.email || ''}" title="${o.email || ''}" onchange="App.updateOrderField('${o.id}', 'email', this.value)">
          </td>
          <td class="col-ord-date">
            <input type="date" class="inline-input" style="font-size:0.82rem;" value="${o.eventDate}" onchange="App.updateOrderField('${o.id}', 'eventDate', this.value)">
          </td>
          <td class="col-ord-time">
            <input type="time" class="inline-input" style="font-size:0.82rem;" value="${o.eventTime || '12:00'}" onchange="App.updateOrderField('${o.id}', 'eventTime', this.value)">
          </td>
          <td class="col-ord-items">
            <input type="text" class="inline-input" value="${o.items}" title="${o.items}" onchange="App.updateOrderField('${o.id}', 'items', this.value)">
          </td>
          <td class="col-ord-status">
            <select class="inline-select" onchange="App.updateOrderField('${o.id}', 'status', this.value)">
              ${['Inquiry', 'Confirmed', 'Baking', 'Decorating', 'Ready for Pickup', 'Completed', 'Cancelled'].map(st => `
                <option value="${st}" ${o.status === st ? 'selected' : ''}>${st}</option>
              `).join('')}
            </select>
          </td>
          <td class="col-ord-del text-right">
            <input type="number" step="0.5" class="inline-input text-right" value="${o.deliveryFee || 0}" onchange="App.updateOrderField('${o.id}', 'deliveryFee', this.value)">
          </td>
          <td class="col-ord-total text-right">
            <input type="number" step="0.5" class="inline-input text-right" style="font-weight:700;" value="${total}" onchange="App.updateOrderField('${o.id}', 'totalQuoted', this.value)">
          </td>
          <td class="col-ord-dep text-right">
            <input type="number" step="0.5" class="inline-input text-right" value="${deposit}" onchange="App.updateOrderField('${o.id}', 'depositPaid', this.value)">
          </td>
          <td class="col-ord-bal text-right" style="font-weight:700; color:${balance > 0 ? 'var(--color-danger)' : 'var(--color-success)'}">$${balance.toFixed(2)}</td>
          <td class="col-ord-notes">
            <input type="text" class="inline-input" style="font-size:0.78rem;" value="${o.notes || ''}" title="${o.notes || ''}" onchange="App.updateOrderField('${o.id}', 'notes', this.value)">
          </td>
          <td class="col-ord-pay"><span class="badge ${payBadgeClass}">${paymentStatus}</span></td>
          <td class="col-ord-days text-center" style="font-family:var(--font-mono); font-size:0.82rem;">${daysLeft}</td>
          <td class="col-ord-act text-center">
            <div style="display:flex; gap:3px; justify-content:center;">
              <button class="btn btn-secondary btn-sm" title="Edit Order" onclick="App.openOrderModal('${o.id}')">✏️</button>
              <button class="btn btn-outline btn-sm" title="Open Invoice" onclick="App.openOrderInInvoice('${o.orderNumber}')">🧾</button>
              <button class="btn btn-danger btn-sm" title="Delete Order" onclick="App.deleteOrder('${o.id}')">🗑️</button>
            </div>
          </td>
        </tr>
      `;
    }).join('');
  }

  updateOrderField(orderId, field, value) {
    const ord = this.data.orders.find(o => o.id === orderId);
    if (!ord) return;

    if (['deliveryFee', 'totalQuoted', 'depositPaid'].includes(field)) {
      ord[field] = parseFloat(value) || 0;
    } else {
      ord[field] = value;
    }

    this.triggerAutoSave();
    this.renderOrders();
    this.renderDashboard();
    this.renderDashboardCharts();
    this.populateInvoiceDropdown();
    this.populateLabelDropdown();
    this.populateContractDropdown();
  }

  openOrderModal(orderId = null) {
    const modal = document.getElementById('orderModal');
    const title = document.getElementById('orderModalTitle');
    
    if (orderId) {
      const o = this.data.orders.find(item => item.id === orderId);
      if (!o) return;
      title.textContent = `Edit Order ${o.orderNumber}`;
      document.getElementById('orderFormId').value = o.id;
      document.getElementById('orderFormNumber').value = o.orderNumber;
      document.getElementById('orderFormName').value = o.customerName;
      document.getElementById('orderFormPhone').value = o.phone || '';
      document.getElementById('orderFormEmail').value = o.email || '';
      document.getElementById('orderFormEventDate').value = o.eventDate;
      document.getElementById('orderFormEventTime').value = o.eventTime || '12:00';
      document.getElementById('orderFormItems').value = o.items;
      document.getElementById('orderFormStatus').value = o.status;
      document.getElementById('orderFormDelivery').value = o.deliveryFee || 0;
      document.getElementById('orderFormTotal').value = o.totalQuoted;
      document.getElementById('orderFormDeposit').value = o.depositPaid;
      document.getElementById('orderFormNotes').value = o.notes || '';
    } else {
      title.textContent = 'Add New Order';
      const nextNum = 'ORD-' + (1000 + this.data.orders.length + 1);
      document.getElementById('orderFormId').value = '';
      document.getElementById('orderFormNumber').value = nextNum;
      document.getElementById('orderFormName').value = '';
      document.getElementById('orderFormPhone').value = '';
      document.getElementById('orderFormEmail').value = '';
      document.getElementById('orderFormEventDate').value = new Date().toISOString().slice(0, 10);
      document.getElementById('orderFormEventTime').value = '12:00';
      document.getElementById('orderFormItems').value = '';
      document.getElementById('orderFormStatus').value = 'Confirmed';
      document.getElementById('orderFormDelivery').value = '15.00';
      document.getElementById('orderFormTotal').value = '150.00';
      document.getElementById('orderFormDeposit').value = '75.00';
      document.getElementById('orderFormNotes').value = '';
    }

    this.calculateModalBalance();
    modal.classList.add('active');
  }

  calculateModalBalance() {
    const total = parseFloat(document.getElementById('orderFormTotal')?.value) || 0;
    const deposit = parseFloat(document.getElementById('orderFormDeposit')?.value) || 0;
    const bal = Math.max(0, total - deposit);
    document.getElementById('orderFormBalance').value = bal.toFixed(2);
    document.getElementById('orderFormPaymentStatus').value = bal <= 0 ? 'Paid in Full' : (deposit > 0 ? 'Partial Deposit' : 'Unpaid');
  }

  handleSaveOrder(e) {
    e.preventDefault();
    const id = document.getElementById('orderFormId').value;
    const orderData = {
      orderNumber: document.getElementById('orderFormNumber').value.trim(),
      customerName: document.getElementById('orderFormName').value.trim(),
      phone: document.getElementById('orderFormPhone').value.trim(),
      email: document.getElementById('orderFormEmail').value.trim(),
      eventDate: document.getElementById('orderFormEventDate').value,
      eventTime: document.getElementById('orderFormEventTime').value,
      items: document.getElementById('orderFormItems').value.trim(),
      status: document.getElementById('orderFormStatus').value,
      deliveryFee: parseFloat(document.getElementById('orderFormDelivery').value) || 0,
      totalQuoted: parseFloat(document.getElementById('orderFormTotal').value) || 0,
      depositPaid: parseFloat(document.getElementById('orderFormDeposit').value) || 0,
      notes: document.getElementById('orderFormNotes').value.trim()
    };

    if (id) {
      const idx = this.data.orders.findIndex(o => o.id === id);
      if (idx !== -1) this.data.orders[idx] = { ...this.data.orders[idx], ...orderData };
    } else {
      orderData.id = 'ord-' + Date.now();
      this.data.orders.push(orderData);
    }

    this.closeModals();
    this.triggerAutoSave();
    this.renderOrders();
    this.renderDashboard();
    this.renderDashboardCharts();
    this.populateInvoiceDropdown();
    this.populateLabelDropdown();
    this.populateContractDropdown();
  }

  deleteOrder(orderId) {
    if (confirm("Delete this order?")) {
      this.data.orders = this.data.orders.filter(o => o.id !== orderId);
      this.triggerAutoSave();
      this.renderOrders();
      this.renderDashboard();
      this.renderDashboardCharts();
    }
  }

  openOrderInInvoice(orderNumber) {
    this.switchView('invoice');
    const select = document.getElementById('invoiceOrderSelect');
    if (select) {
      select.value = orderNumber;
      this.renderInvoice(orderNumber);
    }
  }

  // =================================================================
  // MODULE 3: RECIPE COSTING ENGINE & BATCH SCALER
  // =================================================================
  populateRecipeDropdown() {
    const sel = document.getElementById('recipeSelectDropdown');
    if (!sel) return;
    sel.innerHTML = this.data.recipes.map(r => `
      <option value="${r.id}" ${r.id === this.activeRecipeId ? 'selected' : ''}>${r.name}</option>
    `).join('');
  }

  getActiveRecipe() {
    return this.data.recipes.find(r => r.id === this.activeRecipeId) || this.data.recipes[0];
  }

  renderActiveRecipe() {
    const recipe = this.getActiveRecipe();
    if (!recipe) return;

    document.getElementById('activeRecipeTitle').textContent = `${recipe.name}`;
    document.getElementById('recipeNameInput').value = recipe.name;
    document.getElementById('recipeCategoryInput').value = recipe.category || 'Wedding Tier';
    document.getElementById('recipeYieldInput').value = recipe.yieldServings;
    document.getElementById('recipeLaborHoursInput').value = recipe.laborHours;
    document.getElementById('recipeLaborRateInput').value = recipe.laborRate.toFixed(2);
    document.getElementById('recipeOverheadInput').value = recipe.overheadFactor;
    document.getElementById('recipeMarginInput').value = recipe.targetMargin;

    const units = ['g', 'kg', 'oz', 'lb', 'ml', 'l', 'cup', 'tbsp', 'tsp', 'unit', 'dozen'];

    const tbody = document.getElementById('recipeIngredientsBody');
    if (tbody) {
      tbody.innerHTML = recipe.ingredients.map((ing, idx) => {
        const lineCost = (parseFloat(ing.quantity) || 0) * (parseFloat(ing.unitCost) || 0);
        return `
          <tr>
            <td>
              <input type="text" class="inline-input" value="${ing.name}" list="inventoryDatalist" onchange="App.updateIngredient(${idx}, 'name', this.value)">
            </td>
            <td>
              <input type="number" step="any" class="inline-input text-right" value="${ing.quantity}" onchange="App.updateIngredient(${idx}, 'quantity', this.value)">
            </td>
            <td>
              <select class="inline-select" onchange="App.updateIngredient(${idx}, 'unit', this.value)">
                ${units.map(u => `
                  <option value="${u}" ${ing.unit === u ? 'selected' : ''}>${u}</option>
                `).join('')}
              </select>
            </td>
            <td>
              <input type="number" step="0.0001" class="inline-input text-right" value="${ing.unitCost}" onchange="App.updateIngredient(${idx}, 'unitCost', this.value)">
            </td>
            <td class="text-right" style="font-weight:700; font-family:var(--font-mono);">$${lineCost.toFixed(2)}</td>
            <td class="text-center">
              <button class="btn btn-outline btn-sm" onclick="App.removeIngredientRow(${idx})" title="Remove">✕</button>
            </td>
          </tr>
        `;
      }).join('');
    }

    this.calculateRecipeFinancials();
    this.calculateScaler();
  }

  updateRecipeMetadata(field, value) {
    const recipe = this.getActiveRecipe();
    if (!recipe) return;
    recipe[field] = value;
    this.populateRecipeDropdown();
    this.triggerAutoSave();
  }

  updateIngredient(index, field, value) {
    const recipe = this.getActiveRecipe();
    if (!recipe || !recipe.ingredients[index]) return;

    if (field === 'quantity' || field === 'unitCost') {
      recipe.ingredients[index][field] = parseFloat(value) || 0;
    } else {
      recipe.ingredients[index][field] = value;
      if (field === 'name') {
        const invItem = this.data.inventory.find(i => i.name.toLowerCase() === value.toLowerCase());
        if (invItem) {
          recipe.ingredients[index].unitCost = invItem.baseCost;
          recipe.ingredients[index].unit = invItem.packUnit;
        }
      }
    }

    this.renderActiveRecipe();
    this.triggerAutoSave();
  }

  addIngredientRow() {
    const recipe = this.getActiveRecipe();
    if (!recipe) return;
    recipe.ingredients.push({
      name: 'Flour (All Purpose)',
      quantity: 100,
      unit: 'g',
      unitCost: 0.0025
    });
    this.renderActiveRecipe();
    this.triggerAutoSave();
  }

  removeIngredientRow(index) {
    const recipe = this.getActiveRecipe();
    if (!recipe) return;
    recipe.ingredients.splice(index, 1);
    this.renderActiveRecipe();
    this.triggerAutoSave();
  }

  handleRecipeParamChange() {
    const recipe = this.getActiveRecipe();
    if (!recipe) return;

    recipe.name = document.getElementById('recipeNameInput').value.trim();
    recipe.yieldServings = parseInt(document.getElementById('recipeYieldInput').value) || 12;
    recipe.laborHours = parseFloat(document.getElementById('recipeLaborHoursInput').value) || 1.08333;
    recipe.laborRate = parseFloat(document.getElementById('recipeLaborRateInput').value) || 25.0;
    recipe.overheadFactor = parseFloat(document.getElementById('recipeOverheadInput').value) || 15.0;
    recipe.targetMargin = parseFloat(document.getElementById('recipeMarginInput').value) || 45.0;

    this.calculateRecipeFinancials();
    this.calculateScaler();
    this.triggerAutoSave();
  }

  calculateRecipeFinancials() {
    const recipe = this.getActiveRecipe();
    if (!recipe) return;

    // Direct Materials Subtotal ($8.35)
    const directMaterials = recipe.ingredients.reduce((sum, ing) => {
      return sum + ((parseFloat(ing.quantity) || 0) * (parseFloat(ing.unitCost) || 0));
    }, 0);

    // Direct Labor ($27.08)
    const directLabor = (parseFloat(recipe.laborHours) || 0) * (parseFloat(recipe.laborRate) || 0);

    // Overhead & Utilities ($1.25 -> 15% of direct materials)
    const overhead = directMaterials * ((parseFloat(recipe.overheadFactor) || 0) / 100);

    // Total COGS ($36.68)
    const totalCOGS = directMaterials + directLabor + overhead;
    const costPerServing = recipe.yieldServings > 0 ? (totalCOGS / recipe.yieldServings) : 0;

    // Suggested Retail Price ($66.70)
    const marginRatio = (parseFloat(recipe.targetMargin) || 0) / 100;
    const suggestedBatchPrice = marginRatio < 1 ? (totalCOGS / (1 - marginRatio)) : (totalCOGS * 1.5);
    const suggestedServingPrice = recipe.yieldServings > 0 ? (suggestedBatchPrice / recipe.yieldServings) : 0;
    const grossProfit = suggestedBatchPrice - totalCOGS;
    const foodCostRatio = suggestedBatchPrice > 0 ? ((directMaterials / suggestedBatchPrice) * 100) : 0;

    document.getElementById('sumDirectMaterials').textContent = `$${directMaterials.toFixed(2)}`;
    document.getElementById('sumDirectLabor').textContent = `$${directLabor.toFixed(2)}`;
    document.getElementById('sumOverhead').textContent = `$${overhead.toFixed(2)}`;
    document.getElementById('sumTotalCOGS').textContent = `$${totalCOGS.toFixed(2)}`;
    document.getElementById('sumCostPerServing').textContent = `$${costPerServing.toFixed(2)}`;
    document.getElementById('sumTargetMargin').textContent = `${recipe.targetMargin.toFixed(1)}%`;
    document.getElementById('sumSuggestedBatchPrice').textContent = `$${suggestedBatchPrice.toFixed(2)}`;
    document.getElementById('sumSuggestedServingPrice').textContent = `$${suggestedServingPrice.toFixed(2)}`;
    document.getElementById('sumGrossProfit').textContent = `$${grossProfit.toFixed(2)}`;
    document.getElementById('sumFoodCostRatio').textContent = `${foodCostRatio.toFixed(1)}%`;

    return { directMaterials, directLabor, overhead, totalCOGS, suggestedBatchPrice, grossProfit };
  }

  calculateScaler() {
    const recipe = this.getActiveRecipe();
    if (!recipe) return;

    const multiplier = parseFloat(document.getElementById('batchMultiplierInput')?.value) || 2.0;
    const fin = this.calculateRecipeFinancials();

    const scaledServings = Math.round(recipe.yieldServings * multiplier);
    const scaledMaterials = fin.directMaterials * multiplier;
    const scaledLabor = fin.directLabor * multiplier;
    const scaledCOGS = fin.totalCOGS * multiplier;
    const scaledRevenue = fin.suggestedBatchPrice * multiplier;
    const scaledProfit = scaledRevenue - scaledCOGS;
    const scaledMargin = recipe.targetMargin;
    const scaledPrepHours = (recipe.laborHours * Math.pow(multiplier, 0.767)).toFixed(2);

    document.getElementById('scalerServings').textContent = scaledServings;
    document.getElementById('scalerCOGS').textContent = `$${scaledCOGS.toFixed(2)}`;
    document.getElementById('scalerMaterials').textContent = `$${scaledMaterials.toFixed(2)}`;
    document.getElementById('scalerLabor').textContent = `$${scaledLabor.toFixed(2)}`;
    document.getElementById('scalerRevenue').textContent = `$${scaledRevenue.toFixed(2)}`;
    document.getElementById('scalerProfit').textContent = `$${scaledProfit.toFixed(2)}`;
    document.getElementById('scalerMargin').textContent = `${scaledMargin.toFixed(1)}%`;
    document.getElementById('scalerPrepTime').textContent = `${scaledPrepHours} hrs`;

    const multipliers = [
      { label: '0.5x (Half Batch)', m: 0.5 },
      { label: '1.0x (Standard Batch)', m: 1.0 },
      { label: '2.0x (Double Batch)', m: 2.0 },
      { label: '4.0x (Quad / Event Batch)', m: 4.0 }
    ];

    const matrixBody = document.getElementById('recipeBatchMatrixBody');
    if (matrixBody) {
      matrixBody.innerHTML = multipliers.map(item => {
        const sServings = Math.round(recipe.yieldServings * item.m);
        const sMat = fin.directMaterials * item.m;
        const sLab = fin.directLabor * item.m;
        const sCOGS = fin.totalCOGS * item.m;
        const sPrice = fin.suggestedBatchPrice * item.m;
        const sProf = sPrice - sCOGS;

        return `
          <tr>
            <td><strong>${item.label}</strong></td>
            <td class="text-right">${sServings}</td>
            <td class="text-right">$${sMat.toFixed(2)}</td>
            <td class="text-right">$${sLab.toFixed(2)}</td>
            <td class="text-right" style="font-weight:700;">$${sCOGS.toFixed(2)}</td>
            <td class="text-right" style="color:var(--color-success); font-weight:700;">$${sPrice.toFixed(2)}</td>
            <td class="text-right" style="color:var(--accent-gold); font-weight:700;">$${sProf.toFixed(2)}</td>
          </tr>
        `;
      }).join('');
    }
  }

  duplicateActiveRecipe() {
    const recipe = this.getActiveRecipe();
    if (!recipe) return;
    const dup = JSON.parse(JSON.stringify(recipe));
    dup.id = 'rec-' + Date.now();
    dup.name = `${recipe.name} (Copy)`;
    this.data.recipes.push(dup);
    this.activeRecipeId = dup.id;
    this.populateRecipeDropdown();
    this.renderActiveRecipe();
    this.triggerAutoSave();
  }

  deleteActiveRecipe() {
    if (this.data.recipes.length <= 1) {
      alert("You must keep at least one recipe.");
      return;
    }
    if (confirm(`Delete recipe "${this.getActiveRecipe().name}"?`)) {
      this.data.recipes = this.data.recipes.filter(r => r.id !== this.activeRecipeId);
      this.activeRecipeId = this.data.recipes[0].id;
      this.populateRecipeDropdown();
      this.renderActiveRecipe();
      this.triggerAutoSave();
    }
  }

  openNewRecipeModal() {
    document.getElementById('recipeModal')?.classList.add('active');
  }

  handleCreateRecipe(e) {
    e.preventDefault();
    const name = document.getElementById('newRecipeName').value.trim();
    const category = document.getElementById('newRecipeCategory').value;
    const yieldServings = parseInt(document.getElementById('newRecipeYield').value) || 12;

    const newRec = {
      id: 'rec-' + Date.now(),
      name,
      category,
      yieldServings,
      laborHours: 1.08333,
      laborRate: this.data.settings.laborRate || 25.0,
      overheadFactor: this.data.settings.overheadFactor || 15.0,
      targetMargin: this.data.settings.targetMargin || 45.0,
      ingredients: [
        { name: 'Flour (All Purpose)', quantity: 300, unit: 'g', unitCost: 0.0025 },
        { name: 'Granulated Sugar', quantity: 200, unit: 'g', unitCost: 0.0020 },
        { name: 'Unsalted Butter', quantity: 150, unit: 'g', unitCost: 0.0080 },
        { name: 'Eggs', quantity: 3, unit: 'unit', unitCost: 0.3250 }
      ]
    };

    this.data.recipes.push(newRec);
    this.activeRecipeId = newRec.id;
    this.closeModals();
    this.populateRecipeDropdown();
    this.renderActiveRecipe();
    this.triggerAutoSave();
  }

  printRecipeCard() {
    window.print();
  }

  // =================================================================
  // MODULE 4: MASTER INVENTORY (100% Fit - ZERO Horizontal Scrolling)
  // =================================================================
  renderInventory() {
    const search = document.getElementById('inventorySearchInput')?.value.toLowerCase() || '';
    const categoryFilter = document.getElementById('inventoryCategoryFilter')?.value || 'ALL';
    const lowStockOnly = document.getElementById('inventoryLowStockOnly')?.checked || false;

    const tbody = document.getElementById('inventoryTableBody');
    if (!tbody) return;

    let totalValue = 0;
    let reorderCount = 0;

    const units = ['g', 'kg', 'oz', 'lb', 'ml', 'l', 'cup', 'tbsp', 'tsp', 'unit', 'dozen'];

    const filtered = this.data.inventory.filter(item => {
      const matchSearch = item.name.toLowerCase().includes(search);
      const matchCat = categoryFilter === 'ALL' || item.category === categoryFilter;
      const isLow = item.stockQty <= item.reorderThreshold;
      if (isLow) reorderCount++;
      const val = item.stockQty * item.packPrice;
      totalValue += val;

      if (lowStockOnly && !isLow) return false;
      return matchSearch && matchCat;
    });

    document.getElementById('kpiTotalInventoryValue').textContent = `$${totalValue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById('totInventoryFootValue').textContent = `$${totalValue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    document.getElementById('kpiReorderCount').textContent = reorderCount;
    document.getElementById('kpiTotalSKUs').textContent = this.data.inventory.length;

    tbody.innerHTML = filtered.map(item => {
      const isLow = item.stockQty <= item.reorderThreshold;
      const statusBadge = isLow 
        ? `<span class="badge badge-cancelled" style="font-size:0.7rem; padding:2px 5px;">⚠️ REORDER</span>` 
        : `<span class="badge badge-completed" style="font-size:0.7rem; padding:2px 5px;">✅ In Stock</span>`;
      const valOnHand = item.stockQty * item.packPrice;
      const baseCost = item.packQty > 0 ? (item.packPrice / item.packQty) : 0;

      return `
        <tr>
          <td class="inv-col-name">
            <input type="text" class="inline-input" style="font-weight:600;" value="${item.name}" title="${item.name}" onchange="App.updateInventoryField('${item.id}', 'name', this.value)">
          </td>
          <td class="inv-col-cat">
            <select class="inline-select" style="font-size:0.76rem;" onchange="App.updateInventoryField('${item.id}', 'category', this.value)">
              <option value="Ingredient" ${item.category === 'Ingredient' ? 'selected' : ''}>Ingredient</option>
              <option value="Packaging" ${item.category === 'Packaging' ? 'selected' : ''}>Packaging</option>
              <option value="Sundries" ${item.category === 'Sundries' ? 'selected' : ''}>Sundries</option>
            </select>
          </td>
          <td class="inv-col-packprice text-right">
            <input type="number" step="any" class="inline-input text-right" style="width:60px; padding:3px 4px;" value="${item.packPrice}" onchange="App.updateInventoryField('${item.id}', 'packPrice', this.value)">
          </td>
          <td class="inv-col-packsize text-center">
            <div style="display:inline-flex; align-items:center; justify-content:center; gap:4px;">
              <input type="number" step="any" class="inline-input text-right" style="width:54px; padding:3px 4px; font-weight:600;" value="${item.packQty}" onchange="App.updateInventoryField('${item.id}', 'packQty', this.value)">
              <select class="inline-select" style="min-width:62px; width:auto; padding:3px 4px; font-size:0.78rem;" onchange="App.updateInventoryField('${item.id}', 'packUnit', this.value)">
                ${units.map(u => `
                  <option value="${u}" ${item.packUnit === u ? 'selected' : ''}>${u}</option>
                `).join('')}
              </select>
            </div>
          </td>
          <td class="inv-col-basecost text-right" style="font-family:var(--font-mono); color:var(--text-muted);">$${baseCost.toFixed(4)}</td>
          <td class="inv-col-stock text-center">
            <div class="inv-stepper">
              <button class="inv-stepper-btn" onclick="App.adjustStock('${item.id}', -1)">-</button>
              <input type="number" step="1" class="inline-input text-center" style="width:42px; font-weight:700; padding:2px;" value="${item.stockQty}" onchange="App.updateInventoryField('${item.id}', 'stockQty', this.value)">
              <button class="inv-stepper-btn" onclick="App.adjustStock('${item.id}', 1)">+</button>
            </div>
          </td>
          <td class="inv-col-reorder text-center">
            <input type="number" step="1" class="inline-input text-center" style="width:44px; padding:3px 4px; font-weight:600;" value="${item.reorderThreshold}" onchange="App.updateInventoryField('${item.id}', 'reorderThreshold', this.value)">
          </td>
          <td class="inv-col-status text-center">${statusBadge}</td>
          <td class="inv-col-totalval text-right" style="font-weight:700; font-family:var(--font-mono);">$${valOnHand.toFixed(2)}</td>
          <td class="inv-col-actions text-center">
            <div style="display:flex; gap:2px; justify-content:center;">
              <button class="btn btn-outline btn-sm" style="padding:2px 4px;" title="Duplicate" onclick="App.duplicateInventoryItem('${item.id}')">📋</button>
              <button class="btn btn-danger btn-sm" style="padding:2px 4px;" title="Delete" onclick="App.deleteInventoryItem('${item.id}')">🗑️</button>
            </div>
          </td>
        </tr>
      `;
    }).join('');
  }

  updateInventoryField(itemId, field, value) {
    const item = this.data.inventory.find(i => i.id === itemId);
    if (!item) return;

    if (['packPrice', 'packQty', 'stockQty', 'reorderThreshold'].includes(field)) {
      item[field] = parseFloat(value) || 0;
      if (item.packQty > 0) {
        item.baseCost = item.packPrice / item.packQty;
      }
    } else {
      item[field] = value;
    }

    this.triggerAutoSave();
    this.renderInventory();
    this.renderDashboard();
    this.renderDashboardCharts();
  }

  adjustStock(itemId, delta) {
    const item = this.data.inventory.find(i => i.id === itemId);
    if (item) {
      item.stockQty = Math.max(0, item.stockQty + delta);
      this.triggerAutoSave();
      this.renderInventory();
      this.renderDashboard();
      this.renderDashboardCharts();
    }
  }

  duplicateInventoryItem(itemId) {
    const item = this.data.inventory.find(i => i.id === itemId);
    if (!item) return;
    const dup = JSON.parse(JSON.stringify(item));
    dup.id = 'inv-' + Date.now();
    dup.name = `${item.name} (Copy)`;
    this.data.inventory.push(dup);
    this.triggerAutoSave();
    this.renderInventory();
    this.renderDashboard();
    this.renderDashboardCharts();
  }

  openInventoryModal(itemId = null) {
    const modal = document.getElementById('inventoryModal');
    const title = document.getElementById('inventoryModalTitle');

    if (itemId) {
      const item = this.data.inventory.find(i => i.id === itemId);
      if (!item) return;
      title.textContent = `Edit Inventory Item`;
      document.getElementById('invFormId').value = item.id;
      document.getElementById('invFormName').value = item.name;
      document.getElementById('invFormCategory').value = item.category;
      document.getElementById('invFormUnit').value = item.packUnit;
      document.getElementById('invFormPackPrice').value = item.packPrice;
      document.getElementById('invFormPackQty').value = item.packQty;
      document.getElementById('invFormStockQty').value = item.stockQty;
      document.getElementById('invFormReorderThreshold').value = item.reorderThreshold;
    } else {
      title.textContent = `Add Inventory Item`;
      document.getElementById('invFormId').value = '';
      document.getElementById('invFormName').value = '';
      document.getElementById('invFormCategory').value = 'Ingredient';
      document.getElementById('invFormUnit').value = 'kg';
      document.getElementById('invFormPackPrice').value = '10.00';
      document.getElementById('invFormPackQty').value = '1';
      document.getElementById('invFormStockQty').value = '5';
      document.getElementById('invFormReorderThreshold').value = '2';
    }

    modal.classList.add('active');
  }

  handleSaveInventory(e) {
    e.preventDefault();
    const id = document.getElementById('invFormId').value;
    const packPrice = parseFloat(document.getElementById('invFormPackPrice').value) || 0;
    const packQty = parseFloat(document.getElementById('invFormPackQty').value) || 1;
    const baseCost = packQty > 0 ? (packPrice / packQty) : 0;

    const itemData = {
      name: document.getElementById('invFormName').value.trim(),
      category: document.getElementById('invFormCategory').value,
      packUnit: document.getElementById('invFormUnit').value,
      packPrice,
      packQty,
      baseCost,
      stockQty: parseFloat(document.getElementById('invFormStockQty').value) || 0,
      reorderThreshold: parseFloat(document.getElementById('invFormReorderThreshold').value) || 0
    };

    if (id) {
      const idx = this.data.inventory.findIndex(i => i.id === id);
      if (idx !== -1) this.data.inventory[idx] = { ...this.data.inventory[idx], ...itemData };
    } else {
      itemData.id = 'inv-' + Date.now();
      this.data.inventory.push(itemData);
    }

    this.closeModals();
    this.triggerAutoSave();
    this.renderInventory();
    this.renderDashboard();
    this.renderDashboardCharts();
  }

  deleteInventoryItem(itemId) {
    if (confirm("Delete this inventory item?")) {
      this.data.inventory = this.data.inventory.filter(i => i.id !== itemId);
      this.triggerAutoSave();
      this.renderInventory();
      this.renderDashboard();
      this.renderDashboardCharts();
    }
  }

  // =================================================================
  // MODULE 5: CAKE PORTION & MULTI-TIER PLANNER (Step-by-Step)
  // =================================================================
  renderTierPlanner() {
    const targetGuests = parseInt(document.getElementById('tpTargetGuests')?.value) || 120;
    const portionStandard = document.getElementById('tpPortionStandard')?.value || 'wedding';

    // Populate order dropdown if empty
    const ordSelect = document.getElementById('tpOrderSelect');
    if (ordSelect && ordSelect.children.length <= 1) {
      this.data.orders.forEach(o => {
        const opt = document.createElement('option');
        opt.value = o.id;
        opt.textContent = `${o.orderNumber} - ${o.customerName} (${o.items})`;
        ordSelect.appendChild(opt);
      });
      if (this.data.tierPlanner.orderId) ordSelect.value = this.data.tierPlanner.orderId;
    }

    // Populate Pan Converter
    const convOrig = document.getElementById('convOriginalPan');
    const convTarg = document.getElementById('convTargetPan');
    if (convOrig && convOrig.children.length === 0) {
      this.data.panMatrices.forEach(p => {
        const o1 = document.createElement('option');
        o1.value = p.dimension;
        o1.textContent = `${p.dimension} (${p.shape})`;
        convOrig.appendChild(o1);

        const o2 = document.createElement('option');
        o2.value = p.dimension;
        o2.textContent = `${p.dimension} (${p.shape})`;
        convTarg.appendChild(o2);
      });
      convOrig.value = '8" Round';
      convTarg.value = '10" Round';
      this.calculatePanConverter();
    }

    const tbody = document.getElementById('tpTiersBody');
    if (!tbody) return;

    let activeTiersCount = 0;
    let totalParty = 0;
    let totalWedding = 0;
    let totalDowels = 0;
    let totalBatter = 0;
    let totalHeight = 0;

    tbody.innerHTML = this.data.tierPlanner.tiers.map((t, idx) => {
      if (t.pan) activeTiersCount++;
      totalParty += (parseFloat(t.party) || 0);
      totalWedding += (parseFloat(t.wedding) || 0);
      totalDowels += (parseFloat(t.dowels) || 0);
      totalBatter += (parseFloat(t.batter) || 0);
      totalHeight += (parseFloat(t.height) || 0);

      return `
        <tr>
          <td>
            <input type="text" class="inline-input" style="font-weight:700;" value="${t.level}" onchange="App.updateTierProperty(${idx}, 'level', this.value)">
          </td>
          <td>
            <select class="inline-select" onchange="App.updateTierPan(${idx}, this.value)">
              <option value="">-- None / No Tier --</option>
              ${this.data.panMatrices.map(p => `
                <option value="${p.dimension}" ${t.pan === p.dimension ? 'selected' : ''}>${p.dimension}</option>
              `).join('')}
            </select>
          </td>
          <td class="text-right">
            <input type="number" step="0.5" class="inline-input text-right" value="${t.height}" onchange="App.updateTierProperty(${idx}, 'height', this.value)">
          </td>
          <td class="text-right">
            <input type="number" step="1" class="inline-input text-right" value="${t.party || 0}" onchange="App.updateTierProperty(${idx}, 'party', this.value)">
          </td>
          <td class="text-right">
            <input type="number" step="1" class="inline-input text-right" style="font-weight:700;" value="${t.wedding || 0}" onchange="App.updateTierProperty(${idx}, 'wedding', this.value)">
          </td>
          <td class="text-right">
            <input type="number" step="0.5" class="inline-input text-right" value="${t.batter || 0}" onchange="App.updateTierProperty(${idx}, 'batter', this.value)">
          </td>
          <td class="text-right">
            <input type="number" step="1" class="inline-input text-right" value="${t.dowels || 0}" onchange="App.updateTierProperty(${idx}, 'dowels', this.value)">
          </td>
          <td>
            <input type="text" class="inline-input" style="font-size:0.8rem;" value="${t.board || ''}" onchange="App.updateTierProperty(${idx}, 'board', this.value)">
          </td>
          <td class="text-center">
            <button class="btn btn-outline btn-sm" title="Clear / Delete Tier" onclick="App.deleteTierRow(${idx})">✕</button>
          </td>
        </tr>
      `;
    }).join('');

    const activeServings = portionStandard === 'wedding' ? totalWedding : totalParty;
    const capacityPct = targetGuests > 0 ? ((activeServings / targetGuests) * 100) : 0;
    const buffer = activeServings - targetGuests;

    document.getElementById('tpTotalTiersCount').textContent = activeTiersCount;
    document.getElementById('tpPartyServings').textContent = totalParty;
    document.getElementById('tpWeddingServings').textContent = totalWedding;
    document.getElementById('tpCapacityPct').textContent = `Capacity: ${capacityPct.toFixed(1)}% of ${targetGuests} guests (${portionStandard === 'wedding' ? 'Wedding' : 'Party'} standard)`;
    document.getElementById('tpBufferServings').textContent = buffer >= 0 ? `+${buffer}` : `${buffer}`;
    document.getElementById('tpBufferServings').style.color = buffer >= 0 ? 'var(--color-success)' : 'var(--color-warning)';
    document.getElementById('tpTotalDowels').textContent = totalDowels;
    document.getElementById('tpTotalBatter').textContent = `${totalBatter.toFixed(1)} Cups`;
    document.getElementById('tpTotalHeight').textContent = `Total Cake Height: ${totalHeight.toFixed(1)}"`;

    document.getElementById('totTierHeight').textContent = `${totalHeight.toFixed(1)}"`;
    document.getElementById('totTierParty').textContent = totalParty;
    document.getElementById('totTierWedding').textContent = totalWedding;
    document.getElementById('totTierBatter').textContent = totalBatter.toFixed(1);
    document.getElementById('totTierDowels').textContent = totalDowels;
  }

  updateTierPan(tierIdx, panDimension) {
    const tier = this.data.tierPlanner.tiers[tierIdx];
    if (!tier) return;

    if (!panDimension) {
      tier.pan = '';
      tier.height = 0;
      tier.party = 0;
      tier.wedding = 0;
      tier.batter = 0;
      tier.dowels = 0;
      tier.board = '';
    } else {
      const pan = this.data.panMatrices.find(p => p.dimension === panDimension);
      if (pan) {
        tier.pan = pan.dimension;
        tier.height = tier.height || 4.0;
        tier.party = pan.party;
        tier.wedding = pan.wedding;
        tier.batter = pan.batter;
        tier.dowels = pan.dowels;
        tier.board = pan.board;
      }
    }

    this.renderTierPlanner();
    this.triggerAutoSave();
  }

  updateTierProperty(tierIdx, field, value) {
    const tier = this.data.tierPlanner.tiers[tierIdx];
    if (!tier) return;
    tier[field] = ['height', 'party', 'wedding', 'batter', 'dowels'].includes(field) ? (parseFloat(value) || 0) : value;
    this.renderTierPlanner();
    this.triggerAutoSave();
  }

  addTierRow() {
    const nextIdx = this.data.tierPlanner.tiers.length + 1;
    this.data.tierPlanner.tiers.push({
      id: 'tier-' + Date.now(),
      level: `Tier ${nextIdx}`,
      pan: '6" Round',
      height: 4.0,
      party: 10,
      wedding: 14,
      batter: 4.0,
      dowels: 0,
      board: '8" Cake Board'
    });
    this.renderTierPlanner();
    this.triggerAutoSave();
  }

  deleteTierRow(index) {
    if (this.data.tierPlanner.tiers[index]) {
      this.data.tierPlanner.tiers[index].pan = '';
      this.data.tierPlanner.tiers[index].height = 0;
      this.data.tierPlanner.tiers[index].party = 0;
      this.data.tierPlanner.tiers[index].wedding = 0;
      this.data.tierPlanner.tiers[index].batter = 0;
      this.data.tierPlanner.tiers[index].dowels = 0;
      this.data.tierPlanner.tiers[index].board = '';
      this.renderTierPlanner();
      this.triggerAutoSave();
    }
  }

  calculatePanConverter() {
    const origDim = document.getElementById('convOriginalPan')?.value;
    const targDim = document.getElementById('convTargetPan')?.value;
    const orig = this.data.panMatrices.find(p => p.dimension === origDim);
    const targ = this.data.panMatrices.find(p => p.dimension === targDim);

    if (orig && targ) {
      document.getElementById('convOriginalAreaText').textContent = `Area: ${orig.area} sq in`;
      document.getElementById('convTargetAreaText').textContent = `Area: ${targ.area} sq in`;
      const factor = targ.area / orig.area;
      document.getElementById('convFactorText').textContent = `${factor.toFixed(2)}x`;
      document.getElementById('convGuidanceText').textContent = `Multiply all base recipe ingredients by ${factor.toFixed(2)}x`;
    }
  }

  // =================================================================
  // MODULE 6: WEEKLY PREP & GROCERY AGGREGATOR
  // =================================================================
  renderWeeklyPrep() {
    const startStr = document.getElementById('prepWeekStart')?.value;
    const endStr = document.getElementById('prepWeekEnd')?.value;

    const startDate = startStr ? new Date(startStr) : new Date('2026-10-19');
    const endDate = endStr ? new Date(endStr) : new Date('2026-10-25');

    const pipelineText = document.getElementById('prepActivePipelineText');
    if (pipelineText && startStr && endStr) {
      const sFmt = startDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
      const eFmt = endDate.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
      pipelineText.textContent = `Active Pipeline: Week of ${sFmt} - ${eFmt}`;
    }

    const weekOrders = this.data.orders.filter(o => {
      if (!o.eventDate) return false;
      const d = new Date(o.eventDate);
      return d >= startDate && d <= endDate;
    });

    const ordersBody = document.getElementById('prepScheduledOrdersBody');
    if (ordersBody) {
      ordersBody.innerHTML = weekOrders.map(o => {
        const bal = (parseFloat(o.totalQuoted) || 0) - (parseFloat(o.depositPaid) || 0);
        return `
          <tr>
            <td><strong>${o.orderNumber}</strong></td>
            <td>${o.customerName}</td>
            <td>${o.eventDate}</td>
            <td>${o.eventTime || '12:00'}</td>
            <td>${o.items}</td>
            <td><span class="badge ${this.getStatusBadgeClass(o.status)}">${o.status}</span></td>
            <td class="text-right">$${(parseFloat(o.depositPaid) || 0).toFixed(2)}</td>
            <td class="text-right" style="font-weight:700;">$${bal.toFixed(2)}</td>
          </tr>
        `;
      }).join('');
    }

    const tbody = document.getElementById('prepGroceryTableBody');
    if (!tbody) return;

    let totalRestock = 0;
    const units = ['g', 'kg', 'oz', 'lb', 'ml', 'l', 'cup', 'tbsp', 'tsp', 'unit', 'dozen'];

    tbody.innerHTML = this.data.groceryList.map((item, idx) => {
      const req = parseFloat(item.weeklyReq) || 0;
      const stock = parseFloat(item.stockQty) || 0;
      const shortage = Math.max(0, req - stock);
      
      const invMatch = this.data.inventory.find(i => i.name.toLowerCase() === item.name.toLowerCase());
      const estUnitCost = invMatch ? (invMatch.packPrice / (invMatch.packQty || 1)) : 2.50;
      const restockCost = item.estCost !== undefined ? parseFloat(item.estCost) : (shortage * estUnitCost);
      if (!item.bought) totalRestock += restockCost;

      return `
        <tr style="${item.bought ? 'opacity:0.45; text-decoration:line-through;' : ''}">
          <td class="text-center">
            <input type="checkbox" ${item.bought ? 'checked' : ''} onchange="App.toggleGroceryBought(${idx})">
          </td>
          <td>
            <input type="text" class="inline-input" style="font-weight:600;" value="${item.name}" onchange="App.updateGroceryField(${idx}, 'name', this.value)">
          </td>
          <td>
            <select class="inline-select" onchange="App.updateGroceryField(${idx}, 'category', this.value)">
              <option value="Ingredient" ${item.category === 'Ingredient' ? 'selected' : ''}>Ingredient</option>
              <option value="Packaging" ${item.category === 'Packaging' ? 'selected' : ''}>Packaging</option>
              <option value="Sundries" ${item.category === 'Sundries' ? 'selected' : ''}>Sundries</option>
            </select>
          </td>
          <td>
            <select class="inline-select" onchange="App.updateGroceryField(${idx}, 'unit', this.value)">
              ${units.map(u => `
                <option value="${u}" ${item.unit === u ? 'selected' : ''}>${u}</option>
              `).join('')}
            </select>
          </td>
          <td class="text-right">
            <input type="number" step="any" class="inline-input text-right" value="${req}" onchange="App.updateGroceryField(${idx}, 'weeklyReq', this.value)">
          </td>
          <td class="text-right">
            <input type="number" step="any" class="inline-input text-right" value="${stock}" onchange="App.updateGroceryField(${idx}, 'stockQty', this.value)">
          </td>
          <td class="text-right" style="font-weight:700; color:${shortage > 0 ? 'var(--color-danger)' : 'var(--color-success)'};">${shortage}</td>
          <td class="text-right">
            <input type="number" step="0.5" class="inline-input text-right" style="font-weight:700;" value="${restockCost}" onchange="App.updateGroceryField(${idx}, 'estCost', this.value)">
          </td>
          <td class="text-center">
            <button class="btn btn-outline btn-sm" onclick="App.deleteGroceryRow(${idx})">🗑️</button>
          </td>
        </tr>
      `;
    }).join('');

    document.getElementById('prepTotalRestockCost').textContent = `TOTAL ESTIMATED RESTOCK EXPENSE: $${totalRestock.toFixed(2)}`;
    document.getElementById('totGroceryFootCost').textContent = `$${totalRestock.toFixed(2)}`;
  }

  updateGroceryField(index, field, value) {
    if (this.data.groceryList[index]) {
      this.data.groceryList[index][field] = ['weeklyReq', 'stockQty', 'estCost'].includes(field) ? (parseFloat(value) || 0) : value;
      this.triggerAutoSave();
      this.renderWeeklyPrep();
    }
  }

  toggleGroceryBought(index) {
    if (this.data.groceryList[index]) {
      this.data.groceryList[index].bought = !this.data.groceryList[index].bought;
      this.triggerAutoSave();
      this.renderWeeklyPrep();
    }
  }

  deleteGroceryRow(index) {
    this.data.groceryList.splice(index, 1);
    this.triggerAutoSave();
    this.renderWeeklyPrep();
  }

  openCustomPrepItemModal() {
    document.getElementById('customPrepModal')?.classList.add('active');
  }

  handleSaveCustomPrepItem(e) {
    e.preventDefault();
    const name = document.getElementById('prepFormName').value.trim();
    const category = document.getElementById('prepFormCategory').value;
    const unit = document.getElementById('prepFormUnit').value;
    const weeklyReq = parseFloat(document.getElementById('prepFormRequired').value) || 0;
    const stockQty = parseFloat(document.getElementById('prepFormStock').value) || 0;

    const invMatch = this.data.inventory.find(i => i.name.toLowerCase() === name.toLowerCase());
    const unitCost = invMatch ? (invMatch.packPrice / invMatch.packQty) : 2.50;
    const shortage = Math.max(0, weeklyReq - stockQty);

    this.data.groceryList.push({
      id: 'gr-' + Date.now(),
      name,
      category,
      unit,
      weeklyReq,
      stockQty,
      estCost: shortage * unitCost,
      bought: false
    });

    this.closeModals();
    this.triggerAutoSave();
    this.renderWeeklyPrep();
  }

  // =================================================================
  // MODULE 7: INVOICE GENERATOR (WYSIWYG Editable)
  // =================================================================
  populateInvoiceDropdown() {
    const sel = document.getElementById('invoiceOrderSelect');
    if (!sel) return;
    sel.innerHTML = this.data.orders.map(o => `
      <option value="${o.orderNumber}">${o.orderNumber} - ${o.customerName} ($${o.totalQuoted.toFixed(2)})</option>
    `).join('');

    const targetOrder = this.data.orders.find(o => o.orderNumber === 'ORD-1002') || this.data.orders[0];
    if (targetOrder) {
      sel.value = targetOrder.orderNumber;
      this.renderInvoice(targetOrder.orderNumber);
    }
  }

  renderInvoice(orderNumber) {
    const order = this.data.orders.find(o => o.orderNumber === orderNumber) || this.data.orders[0];
    if (!order) return;

    const s = this.data.settings;
    document.getElementById('invBakeryName').textContent = s.bakeryName;
    document.getElementById('invBakeryAddress').textContent = s.address;
    document.getElementById('invBakeryContact').textContent = `Phone: ${s.phone} | Email: ${s.email}`;

    document.getElementById('invNumber').textContent = order.orderNumber;
    document.getElementById('invDate').textContent = '10/16/2026';

    document.getElementById('invClientName').textContent = order.customerName;
    document.getElementById('invClientPhone').textContent = order.phone || '-';
    document.getElementById('invClientEmail').textContent = order.email || '-';

    document.getElementById('invEventDate').textContent = order.eventDate;
    document.getElementById('invEventTime').textContent = order.eventTime || '11:00';
    document.getElementById('invStatus').textContent = order.status;

    if (!this.data.invoiceCustomItems[order.orderNumber]) {
      const itemQuoted = (parseFloat(order.totalQuoted) || 0) - (parseFloat(order.deliveryFee) || 0);
      this.data.invoiceCustomItems[order.orderNumber] = [
        { description: order.items, category: 'Item Quoted', amount: itemQuoted },
        { description: 'Delivery & Setup Fee', category: 'Service', amount: parseFloat(order.deliveryFee) || 0 }
      ];
    }

    const items = this.data.invoiceCustomItems[order.orderNumber];
    const tbody = document.getElementById('invItemsBody');
    if (tbody) {
      tbody.innerHTML = items.map((it, idx) => `
        <tr>
          <td contenteditable="true" onblur="App.updateInvoiceItem(${idx}, 'description', this.textContent)">${it.description}</td>
          <td contenteditable="true" onblur="App.updateInvoiceItem(${idx}, 'category', this.textContent)">${it.category}</td>
          <td style="text-align:right;" contenteditable="true" onblur="App.updateInvoiceItem(${idx}, 'amount', this.textContent)">$${(parseFloat(it.amount) || 0).toFixed(2)}</td>
          <td class="no-print text-center">
            <button class="btn btn-outline btn-sm" onclick="App.deleteInvoiceItem(${idx})">✕</button>
          </td>
        </tr>
      `).join('');
    }

    document.getElementById('invNotes').textContent = order.notes || "Pickup at bakery, 'Happy 30th Mike'";

    const total = parseFloat(order.totalQuoted) || 0;
    const deposit = parseFloat(order.depositPaid) || 0;
    const balance = Math.max(0, total - deposit);

    document.getElementById('invTotalQuoted').textContent = `$${total.toFixed(2)}`;
    document.getElementById('invDepositPaid').textContent = `$${deposit.toFixed(2)}`;
    document.getElementById('invBalanceDue').textContent = `$${balance.toFixed(2)}`;

    document.getElementById('invBankingDetails').innerHTML = s.bankingInfo.replace(/\n/g, '<br>');

    this.bindWysiwygSave('invBakeryName', 'bakeryName');
    this.bindWysiwygSave('invBakeryAddress', 'address');
    this.bindWysiwygSave('invBankingDetails', 'bankingInfo');
  }

  updateInvoiceItem(idx, field, val) {
    const orderNum = document.getElementById('invoiceOrderSelect')?.value;
    const items = this.data.invoiceCustomItems[orderNum];
    if (items && items[idx]) {
      if (field === 'amount') {
        const num = parseFloat(val.replace(/[$,]/g, '')) || 0;
        items[idx].amount = num;
      } else {
        items[idx][field] = val.trim();
      }
      this.triggerAutoSave();
    }
  }

  addInvoiceLineItem() {
    const orderNum = document.getElementById('invoiceOrderSelect')?.value;
    if (!this.data.invoiceCustomItems[orderNum]) {
      this.data.invoiceCustomItems[orderNum] = [];
    }
    this.data.invoiceCustomItems[orderNum].push({
      description: 'Custom Cake Stand Rental / Decor',
      category: 'Sundries',
      amount: 25.00
    });
    this.triggerAutoSave();
    this.renderInvoice(orderNum);
  }

  deleteInvoiceItem(idx) {
    const orderNum = document.getElementById('invoiceOrderSelect')?.value;
    if (this.data.invoiceCustomItems[orderNum]) {
      this.data.invoiceCustomItems[orderNum].splice(idx, 1);
      this.triggerAutoSave();
      this.renderInvoice(orderNum);
    }
  }

  bindWysiwygSave(elementId, settingsKey) {
    const el = document.getElementById(elementId);
    if (el) {
      el.onblur = () => {
        this.data.settings[settingsKey] = el.innerText;
        this.triggerAutoSave();
      };
    }
  }

  // =================================================================
  // MODULE 8: COTTAGE FOOD BOX LABELS
  // =================================================================
  populateLabelDropdown() {
    const sel = document.getElementById('labelOrderSelect');
    if (!sel) return;
    sel.innerHTML = this.data.orders.map(o => `
      <option value="${o.orderNumber}">${o.orderNumber} - ${o.customerName}</option>
    `).join('');

    const target = this.data.orders.find(o => o.orderNumber === 'ORD-1005') || this.data.orders[0];
    if (target) {
      sel.value = target.orderNumber;
      this.renderCottageLabel(target.orderNumber);
    }
  }

  renderCottageLabel(orderNumber) {
    const order = this.data.orders.find(o => o.orderNumber === orderNumber) || this.data.orders[0];
    if (!order) return;

    const s = this.data.settings;
    document.getElementById('lblBakeryTitle').textContent = s.bakeryName.toUpperCase();
    document.getElementById('lblBakerySub').textContent = `${s.address} • Phone: ${s.phone}`;
    document.getElementById('lblOrderId').textContent = order.orderNumber;
    document.getElementById('lblClient').textContent = order.customerName;
    document.getElementById('lblFulfillmentDate').textContent = '2026-10-28';
    document.getElementById('lblProduct').textContent = order.items;
    document.getElementById('labelProductInput').value = order.items;

    document.getElementById('lblPermit').textContent = s.cfoLicense;
    document.getElementById('lblDisclosure').textContent = s.stateDisclosure;
    document.getElementById('lblSafeHandling').textContent = s.safeHandling;
    document.getElementById('lblBakerSign').textContent = `Made with love by ${s.chefName} • ${s.bakeryName}`;

    const syncField = (inputId, displayId, settingKey = null) => {
      const inp = document.getElementById(inputId);
      const disp = document.getElementById(displayId);
      if (inp && disp) {
        inp.oninput = () => {
          disp.textContent = inp.value;
          if (settingKey) this.data.settings[settingKey] = inp.value;
          this.triggerAutoSave();
        };
        disp.onblur = () => {
          inp.value = disp.innerText;
          if (settingKey) this.data.settings[settingKey] = disp.innerText;
          this.triggerAutoSave();
        };
      }
    };

    syncField('labelProductInput', 'lblProduct');
    syncField('labelWeightInput', 'lblWeight');
    syncField('labelFulfillmentInput', 'lblFulfillmentDate');
    syncField('labelAllergensInput', 'lblAllergens');
    syncField('labelIngredientsInput', 'lblIngredients');
    syncField('labelDisclosureInput', 'lblDisclosure', 'stateDisclosure');
    syncField('labelBestByInput', 'lblBestBy');
    syncField('labelPermitInput', 'lblPermit', 'cfoLicense');
  }

  // =================================================================
  // MODULE 9: BAKERY CONTRACT AGREEMENT
  // =================================================================
  populateContractDropdown() {
    const sel = document.getElementById('contractOrderSelect');
    if (!sel) return;
    sel.innerHTML = this.data.orders.map(o => `
      <option value="${o.orderNumber}">${o.orderNumber} - ${o.customerName} (${o.eventDate})</option>
    `).join('');

    const target = this.data.orders.find(o => o.orderNumber === 'ORD-1001') || this.data.orders[0];
    if (target) {
      sel.value = target.orderNumber;
      this.renderContract(target.orderNumber);
    }
  }

  renderContract(orderNumber) {
    const order = this.data.orders.find(o => o.orderNumber === orderNumber) || this.data.orders[0];
    if (!order) return;

    const s = this.data.settings;
    document.getElementById('cntBakeryName').textContent = s.bakeryName.toUpperCase();
    document.getElementById('cntBakeryContact').textContent = `${s.address} • Phone: ${s.phone} • Email: ${s.email}`;

    document.getElementById('cntOrderId').textContent = order.orderNumber;
    document.getElementById('cntClientName').textContent = order.customerName;
    document.getElementById('cntEventDate').textContent = order.eventDate;
    document.getElementById('cntOrderItems').textContent = order.items;

    const total = parseFloat(order.totalQuoted) || 0;
    const deposit = parseFloat(order.depositPaid) || 0;
    const balance = Math.max(0, total - deposit);

    document.getElementById('cntTotalQuoted').textContent = `$${total.toFixed(2)}`;
    document.getElementById('cntBalanceDue').textContent = `$${balance.toFixed(2)}`;

    document.getElementById('cntSigClientName').textContent = order.customerName;
    document.getElementById('cntSigBakerName').textContent = `${s.chefName}, ${s.bakeryName}`;

    const container = document.getElementById('contractClausesContainer');
    if (container) {
      container.innerHTML = this.data.contractTerms.map((term, idx) => `
        <div style="position:relative; padding-right:30px;">
          <p>
            <strong contenteditable="true" onblur="App.updateContractClause(${idx}, 'title', this.innerText)">${term.title}</strong>
            <span contenteditable="true" onblur="App.updateContractClause(${idx}, 'body', this.innerText)"> ${term.body}</span>
          </p>
          <button class="no-print btn btn-outline btn-sm" style="position:absolute; right:0; top:0;" onclick="App.deleteContractClause(${idx})">✕</button>
        </div>
      `).join('');
    }
  }

  updateContractClause(index, field, value) {
    if (this.data.contractTerms[index]) {
      this.data.contractTerms[index][field] = value.trim();
      this.triggerAutoSave();
    }
  }

  addContractClause() {
    const num = this.data.contractTerms.length + 1;
    this.data.contractTerms.push({
      id: 'ct-' + Date.now(),
      title: `${num}. Custom Clause Title:`,
      body: 'Insert your custom legal policy or dietary requirements here.'
    });
    this.triggerAutoSave();
    const curOrd = document.getElementById('contractOrderSelect')?.value;
    this.renderContract(curOrd);
  }

  deleteContractClause(index) {
    this.data.contractTerms.splice(index, 1);
    this.triggerAutoSave();
    const curOrd = document.getElementById('contractOrderSelect')?.value;
    this.renderContract(curOrd);
  }

  resetDefaultContract() {
    if (confirm("Reset contract clauses to standard BakeSuite agreement?")) {
      this.data.contractTerms = JSON.parse(JSON.stringify(DEFAULT_BAKERY_DATA.contractTerms));
      this.triggerAutoSave();
      const curOrd = document.getElementById('contractOrderSelect')?.value;
      this.renderContract(curOrd);
    }
  }

  // =================================================================
  // MODULE 10: SETTINGS & MATRICES
  // =================================================================
  renderSettings() {
    const s = this.data.settings;
    document.getElementById('cfgBakeryName').value = s.bakeryName;
    document.getElementById('cfgChefName').value = s.chefName;
    document.getElementById('cfgPhone').value = s.phone;
    document.getElementById('cfgEmail').value = s.email;
    document.getElementById('cfgAddress').value = s.address;
    document.getElementById('cfgCurrency').value = s.currency;
    document.getElementById('cfgLaborRate').value = s.laborRate;
    document.getElementById('cfgOverhead').value = s.overheadFactor;
    document.getElementById('cfgMargin').value = s.targetMargin;
    document.getElementById('cfgDeliveryBase').value = s.deliveryBaseFee;
    document.getElementById('cfgMileageRate').value = s.mileageRate;
    document.getElementById('cfgDepositReq').value = s.depositReqPercent;
    document.getElementById('cfgTaxRate').value = s.salesTaxRate;
    document.getElementById('cfgCfoLicense').value = s.cfoLicense;
    document.getElementById('cfgBankingInfo').value = s.bankingInfo;

    const matrixBody = document.getElementById('cfgPanMatrixBody');
    if (matrixBody) {
      matrixBody.innerHTML = this.data.panMatrices.map((p, idx) => `
        <tr>
          <td>
            <input type="text" class="inline-input" style="font-weight:600; width:100%;" value="${p.dimension}" onchange="App.updatePanMatrixField(${idx}, 'dimension', this.value)">
          </td>
          <td style="min-width:115px;">
            <select class="inline-select" style="min-width:105px; width:100%; font-weight:600; padding:4px 6px;" onchange="App.updatePanMatrixField(${idx}, 'shape', this.value)">
              <option value="Round" ${p.shape === 'Round' ? 'selected' : ''}>Round</option>
              <option value="Square" ${p.shape === 'Square' ? 'selected' : ''}>Square</option>
              <option value="Heart" ${p.shape === 'Heart' ? 'selected' : ''}>Heart</option>
              <option value="Hexagon" ${p.shape === 'Hexagon' ? 'selected' : ''}>Hexagon</option>
              <option value="Sheet" ${p.shape === 'Sheet' ? 'selected' : ''}>Sheet</option>
            </select>
          </td>
          <td class="text-right">
            <input type="number" step="1" class="inline-input text-right" value="${p.party}" onchange="App.updatePanMatrixField(${idx}, 'party', this.value)">
          </td>
          <td class="text-right">
            <input type="number" step="1" class="inline-input text-right" style="font-weight:700;" value="${p.wedding}" onchange="App.updatePanMatrixField(${idx}, 'wedding', this.value)">
          </td>
          <td class="text-right">
            <input type="number" step="0.1" class="inline-input text-right" value="${p.area}" onchange="App.updatePanMatrixField(${idx}, 'area', this.value)">
          </td>
          <td class="text-center">
            <button class="btn btn-outline btn-sm" onclick="App.deletePanMatrixRow(${idx})">🗑️</button>
          </td>
        </tr>
      `).join('');
    }

    const unitBody = document.getElementById('cfgUnitMatrixBody');
    if (unitBody) {
      unitBody.innerHTML = this.data.unitConversions.map((u, idx) => `
        <tr>
          <td>
            <input type="text" class="inline-input" style="font-weight:700;" value="${u.unit}" onchange="App.updateUnitConversionField(${idx}, 'unit', this.value)">
          </td>
          <td>
            <select class="inline-select" onchange="App.updateUnitConversionField(${idx}, 'category', this.value)">
              <option value="Mass" ${u.category === 'Mass' ? 'selected' : ''}>Mass</option>
              <option value="Volume" ${u.category === 'Volume' ? 'selected' : ''}>Volume</option>
              <option value="Count" ${u.category === 'Count' ? 'selected' : ''}>Count</option>
            </select>
          </td>
          <td>
            <input type="text" class="inline-input" value="${u.baseUnit}" onchange="App.updateUnitConversionField(${idx}, 'baseUnit', this.value)">
          </td>
          <td class="text-right">
            <input type="number" step="any" class="inline-input text-right" style="font-family:var(--font-mono);" value="${u.multiplier}" onchange="App.updateUnitConversionField(${idx}, 'multiplier', this.value)">
          </td>
          <td class="text-center">
            <button class="btn btn-outline btn-sm" onclick="App.deleteUnitConversionRow(${idx})">🗑️</button>
          </td>
        </tr>
      `).join('');
    }

    this.renderLicenseStatusInSettings();
  }

  updatePanMatrixField(index, field, value) {
    if (this.data.panMatrices[index]) {
      this.data.panMatrices[index][field] = ['party', 'wedding', 'area', 'batter', 'dowels'].includes(field) ? (parseFloat(value) || 0) : value;
      this.triggerAutoSave();
    }
  }

  addPanMatrixRow() {
    this.data.panMatrices.push({
      id: 'pan-' + Date.now(),
      dimension: '16" Round',
      shape: 'Round',
      party: 85,
      wedding: 110,
      area: 201.1,
      batter: 30.0,
      dowels: 12,
      board: '18" Cake Board'
    });
    this.triggerAutoSave();
    this.renderSettings();
  }

  deletePanMatrixRow(index) {
    this.data.panMatrices.splice(index, 1);
    this.triggerAutoSave();
    this.renderSettings();
  }

  updateUnitConversionField(index, field, value) {
    if (this.data.unitConversions[index]) {
      this.data.unitConversions[index][field] = field === 'multiplier' ? (parseFloat(value) || 1.0) : value;
      this.triggerAutoSave();
    }
  }

  addUnitConversionRow() {
    this.data.unitConversions.push({
      id: 'u-' + Date.now(),
      unit: 'pinch',
      category: 'Mass',
      baseUnit: 'g',
      multiplier: 0.36
    });
    this.triggerAutoSave();
    this.renderSettings();
  }

  deleteUnitConversionRow(index) {
    this.data.unitConversions.splice(index, 1);
    this.triggerAutoSave();
    this.renderSettings();
  }

  saveSettings() {
    const s = this.data.settings;
    s.bakeryName = document.getElementById('cfgBakeryName').value.trim();
    s.chefName = document.getElementById('cfgChefName').value.trim();
    s.phone = document.getElementById('cfgPhone').value.trim();
    s.email = document.getElementById('cfgEmail').value.trim();
    s.address = document.getElementById('cfgAddress').value.trim();
    s.currency = document.getElementById('cfgCurrency').value.trim();
    s.laborRate = parseFloat(document.getElementById('cfgLaborRate').value) || 25.0;
    s.overheadFactor = parseFloat(document.getElementById('cfgOverhead').value) || 15.0;
    s.targetMargin = parseFloat(document.getElementById('cfgMargin').value) || 45.0;
    s.deliveryBaseFee = parseFloat(document.getElementById('cfgDeliveryBase').value) || 15.0;
    s.mileageRate = parseFloat(document.getElementById('cfgMileageRate').value) || 1.5;
    s.depositReqPercent = parseFloat(document.getElementById('cfgDepositReq').value) || 50.0;
    s.salesTaxRate = parseFloat(document.getElementById('cfgTaxRate').value) || 6.5;
    s.cfoLicense = document.getElementById('cfgCfoLicense').value.trim();
    s.bankingInfo = document.getElementById('cfgBankingInfo').value.trim();

    this.triggerAutoSave();
    this.renderAllViews();
    alert("Configuration and business profile successfully saved!");
  }

  // =================================================================
  // MODULE 0: ONE-TIME LICENSE ACTIVATION & OFFLINE VERIFICATION
  // =================================================================
  async checkLicenseState() {
    try {
      // 1. Check LocalStorage
      let savedLicense = null;
      const lsLicense = localStorage.getItem('bakesuite_license_v1');
      if (lsLicense) {
        try { savedLicense = JSON.parse(lsLicense); } catch (e) {}
      }

      // 2. Check IndexedDB if not in LocalStorage
      if (!savedLicense || !savedLicense.activated) {
        savedLicense = await this.idbGet('bakesuite_license');
      }

      // 3. Check master bakery data file
      if ((!savedLicense || !savedLicense.activated) && this.data && this.data.license && this.data.license.activated) {
        savedLicense = this.data.license;
      }

      if (savedLicense && savedLicense.activated) {
        this.license = savedLicense;
        localStorage.setItem('bakesuite_license_v1', JSON.stringify(savedLicense));
        await this.idbSet('bakesuite_license', savedLicense);
        if (this.data && !this.data.license) {
          this.data.license = savedLicense;
        }
        this.unlockApp(false);
        this.renderLicenseStatusInSettings();
      } else {
        this.lockApp();
      }
    } catch (err) {
      console.warn("License state check error:", err);
      this.lockApp();
    }
  }

  unlockApp(animate = false) {
    const overlay = document.getElementById('licenseLockOverlay');
    document.body.classList.remove('license-locked');
    if (overlay) {
      if (animate) {
        overlay.style.transition = 'opacity 0.35s ease, visibility 0.35s ease';
        overlay.style.opacity = '0';
        setTimeout(() => {
          overlay.classList.remove('active');
          overlay.style.display = 'none';
        }, 350);
      } else {
        overlay.classList.remove('active');
        overlay.style.display = 'none';
      }
    }
    const errEl = document.getElementById('licenseErrorMsg');
    if (errEl) errEl.style.display = 'none';
  }

  lockApp() {
    document.body.classList.add('license-locked');
    const overlay = document.getElementById('licenseLockOverlay');
    if (overlay) {
      overlay.style.transition = '';
      overlay.style.opacity = '1';
      overlay.style.display = 'flex';
      overlay.classList.add('active');
    }
    setTimeout(() => {
      document.getElementById('licenseKeyInput')?.focus();
    }, 100);
  }

  sha256(ascii) {
    function rightRotate(value, amount) {
      return (value >>> amount) | (value << (32 - amount));
    }
    const mathPow = Math.pow;
    const maxWord = mathPow(2, 32);
    let lengthProperty = 'length';
    let i, j;
    let result = '';

    const words = [];
    const asciiBitLength = ascii[lengthProperty] * 8;
    
    let hash = [];
    let k = [];
    let primeCounter = 0;

    const isComposite = {};
    for (let candidate = 2; primeCounter < 64; candidate++) {
      if (!isComposite[candidate]) {
        for (i = candidate * 2; i < 313; i += candidate) {
          isComposite[i] = true;
        }
        if (primeCounter < 8) {
          hash[primeCounter] = (mathPow(candidate, 0.5) * maxWord) | 0;
        }
        k[primeCounter] = (mathPow(candidate, 1 / 3) * maxWord) | 0;
        primeCounter++;
      }
    }

    ascii += '\x80';
    while (ascii[lengthProperty] % 64 - 56) ascii += '\x00';
    for (i = 0; i < ascii[lengthProperty]; i++) {
      j = ascii.charCodeAt(i);
      words[i >> 2] |= j << ((3 - i % 4) * 8);
    }
    words[words[lengthProperty]] = ((asciiBitLength / maxWord) | 0);
    words[words[lengthProperty]] = (asciiBitLength | 0);

    for (j = 0; j < words[lengthProperty];) {
      const w = words.slice(j, j += 16);
      const oldHash = hash;
      hash = hash.slice(0, 8);

      for (i = 0; i < 64; i++) {
        const i2 = i + j;
        const w15 = w[i - 15], w2 = w[i - 2];

        const a = hash[0], e = hash[4];
        const temp1 = hash[7]
          + (rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25))
          + ((e & hash[5]) ^ ((~e) & hash[6]))
          + k[i]
          + (w[i] = (i < 16) ? w[i] : (
              w[i - 16]
              + (rightRotate(w15, 7) ^ rightRotate(w15, 18) ^ (w15 >>> 3))
              + w[i - 7]
              + (rightRotate(w2, 17) ^ rightRotate(w2, 19) ^ (w2 >>> 10))
            ) | 0
          );
        const temp2 = (rightRotate(a, 2) ^ rightRotate(a, 13) ^ rightRotate(a, 22))
          + ((a & hash[1]) ^ (a & hash[2]) ^ (hash[1] & hash[2]));

        hash = [(temp1 + temp2) | 0].concat(hash);
        hash[4] = (hash[4] + temp1) | 0;
      }

      for (i = 0; i < 8; i++) {
        hash[i] = (hash[i] + oldHash[i]) | 0;
      }
    }

    for (i = 0; i < 8; i++) {
      for (j = 3; j >= 0; j--) {
        const b = (hash[i] >> (j * 8)) & 255;
        result += ((b < 16) ? '0' : '') + b.toString(16);
      }
    }
    return result;
  }

  computeLicenseChecksum(p1, p2) {
    const salt = 'BakeSuite_OS_Secret_2026';
    const str = `${p1}-${p2}-${salt}`;
    let hash = 2166136261;
    for (let i = 0; i < str.length; i++) {
      hash ^= str.charCodeAt(i);
      hash = Math.imul(hash, 16777619);
    }
    const hex = (hash >>> 0).toString(16).toUpperCase().padStart(8, '0');
    return hex.slice(-4);
  }

  validateLicenseKey(rawKey) {
    if (!rawKey || typeof rawKey !== 'string') return null;
    const key = rawKey.trim().toUpperCase();

    // Method A: Obfuscated Master Buyer & Admin Hashes
    const masterHashes = [
      '957f998caa0cd16f5bf7208f6461c7edd50a97c1122cf113c8d48e009cfdf50f',
      '1e3796bd38284ea172b85428cfcbe7a9685fdefda27f5b7647cc6af2d4187233'
    ];
    const inputHash = this.sha256(key);
    if (masterHashes.includes(inputHash)) {
      const last4 = key.slice(-4);
      return {
        type: 'master',
        keyMask: `BSPRO-****-****-${last4}`,
        activatedAt: new Date().toISOString()
      };
    }

    // Method B: Algorithmic Checksum for unique BSPRO-XXXX-XXXX-CCCC keys
    const parts = key.split('-');
    if (parts.length === 4 && parts[0] === 'BSPRO' && parts[1].length === 4 && parts[2].length === 4 && parts[3].length === 4) {
      const expectedChecksum = this.computeLicenseChecksum(parts[1], parts[2]);
      if (parts[3] === expectedChecksum) {
        return {
          type: 'algorithmic',
          keyMask: `BSPRO-****-****-${parts[3]}`,
          activatedAt: new Date().toISOString()
        };
      }
    }

    return null;
  }

  verifyAndActivateLicense() {
    const inputEl = document.getElementById('licenseKeyInput');
    const errEl = document.getElementById('licenseErrorMsg');
    if (!inputEl) return;
    const rawKey = inputEl.value.trim().toUpperCase();
    inputEl.value = rawKey;

    const validInfo = this.validateLicenseKey(rawKey);
    if (!validInfo) {
      if (errEl) {
        errEl.textContent = '❌ Invalid license key. Please check your purchase confirmation or PDF guide.';
        errEl.style.display = 'block';
      }
      inputEl.focus();
      return;
    }

    if (errEl) errEl.style.display = 'none';

    const licenseData = {
      activated: true,
      keyMask: validInfo.keyMask,
      activatedAt: validInfo.activatedAt,
      type: validInfo.type
    };

    this.saveLicense(licenseData);
    this.unlockApp(true);
  }

  saveLicense(licenseData) {
    this.license = licenseData;
    localStorage.setItem('bakesuite_license_v1', JSON.stringify(licenseData));
    this.idbSet('bakesuite_license', licenseData);
    if (this.data) {
      this.data.license = licenseData;
      this.triggerAutoSave();
    }
    this.renderLicenseStatusInSettings();
  }

  deactivateLicense() {
    if (confirm("Are you sure you want to deactivate your license on this device? BakeSuite Pro will lock until a valid key is entered.")) {
      localStorage.removeItem('bakesuite_license_v1');
      this.idbSet('bakesuite_license', null);
      if (this.data) {
        this.data.license = null;
        this.triggerAutoSave();
      }
      this.license = null;
      this.lockApp();
      const inputEl = document.getElementById('licenseKeyInput');
      if (inputEl) inputEl.value = '';
      const errEl = document.getElementById('licenseErrorMsg');
      if (errEl) errEl.style.display = 'none';
      this.renderLicenseStatusInSettings();
    }
  }

  async restoreLicenseViaOpenFile() {
    if ('showOpenFilePicker' in window) {
      try {
        const [handle] = await window.showOpenFilePicker({
          types: [{ description: 'JSON Document', accept: { 'application/json': ['.json'] } }],
          multiple: false
        });
        const file = await handle.getFile();
        const text = await file.text();
        const parsed = JSON.parse(text);

        if (!parsed || !parsed.orders || !parsed.inventory) {
          alert("The selected file is not a valid BakeSuite Pro master database.");
          return;
        }

        if (parsed.license && parsed.license.activated) {
          this.data = this.mergeWithDefaultSchema(parsed);
          this.fileHandle = handle;
          await this.idbSet('masterFileHandle', handle);
          await this.idbSet('bakeryMasterData', this.data);
          this.saveLicense(parsed.license);
          this.renderAllViews();
          this.unlockApp(true);
          alert(`Welcome back! License verified from master backup (${parsed.license.keyMask}). BakeSuite Pro unlocked.`);
        } else {
          alert("This backup file does not contain an active license. Please enter your license key to unlock.");
        }
      } catch (err) {
        if (err.name !== 'AbortError') {
          document.getElementById('licenseBackupUpload')?.click();
        }
      }
    } else {
      document.getElementById('licenseBackupUpload')?.click();
    }
  }

  handleLicenseBackupImport(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const parsed = JSON.parse(e.target.result);
        if (!parsed || !parsed.orders || !parsed.inventory) {
          alert("The selected file is not a valid BakeSuite Pro master database.");
          return;
        }

        if (parsed.license && parsed.license.activated) {
          this.data = this.mergeWithDefaultSchema(parsed);
          await this.idbSet('bakeryMasterData', this.data);
          this.saveLicense(parsed.license);
          this.renderAllViews();
          this.unlockApp(true);
          alert(`Welcome back! License verified from master backup (${parsed.license.keyMask}). BakeSuite Pro unlocked.`);
        } else {
          alert("This backup file does not contain an active license. Please enter your license key to unlock.");
        }
      } catch (err) {
        alert("Error reading backup file: " + err.message);
      }
    };
    reader.readAsText(file);
    event.target.value = '';
  }

  renderLicenseStatusInSettings() {
    const badgeEl = document.getElementById('cfgLicenseBadge');
    const maskEl = document.getElementById('cfgLicenseKeyMask');
    const dateEl = document.getElementById('cfgLicenseDate');

    if (this.license && this.license.activated) {
      if (badgeEl) {
        badgeEl.className = 'badge badge-completed';
        badgeEl.textContent = '✓ Active Commercial License';
      }
      if (maskEl) maskEl.textContent = this.license.keyMask || 'BSPRO-****-****';
      if (dateEl) {
        dateEl.textContent = this.license.activatedAt ? new Date(this.license.activatedAt).toLocaleDateString() : 'Active';
      }
    } else {
      if (badgeEl) {
        badgeEl.className = 'badge badge-cancelled';
        badgeEl.textContent = '⚠ Unactivated';
      }
      if (maskEl) maskEl.textContent = 'Not Activated';
      if (dateEl) dateEl.textContent = 'None';
    }
  }

  generateLicenseKey(count = 1) {
    const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    const keys = [];
    for (let c = 0; c < count; c++) {
      let p1 = '';
      let p2 = '';
      for (let i = 0; i < 4; i++) {
        p1 += chars.charAt(Math.floor(Math.random() * chars.length));
        p2 += chars.charAt(Math.floor(Math.random() * chars.length));
      }
      const chk = this.computeLicenseChecksum(p1, p2);
      keys.push(`BSPRO-${p1}-${p2}-${chk}`);
    }
    console.log("%c🔑 BakeSuite Pro Unique License Key Generator", "color: #10b981; font-weight: bold; font-size: 14px;");
    keys.forEach((k, idx) => console.log(`${idx + 1}. %c${k}`, "color: #f59e0b; font-family: monospace; font-size: 13px;"));
    return count === 1 ? keys[0] : keys;
  }

  closeModals() {
    document.querySelectorAll('.modal-overlay').forEach(m => m.classList.remove('active'));
  }
}

// Global App Instance & Initialization
const App = new BakeryApp();
window.addEventListener('DOMContentLoaded', () => {
  App.init();
});
