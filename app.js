(() => {
  "use strict";

  const APP_CONFIG = window.APP_CONFIG || {};
  const FIXED_APPS_SCRIPT_BASE_URL =
    "https://script.google.com/macros/s/AKfycbxEkgzuv0id6ahjSORd87458qTVSlFpVAw4Yixea3PCa3c90DcZN3iqNF4LmIkLnarF/exec";

  const STORAGE_KEYS = {
    session: "dorm-market-session-v2",
    activeTab: "dorm-market-active-tab-v2",
    listingsCache: "dorm-market-listings-cache-v1",
  };

  const SUBJECT_OPTIONS = [
    "数学",
    "物理",
    "化学",
    "情報",
    "電気電子",
    "機械",
    "建築",
    "材料",
    "英語",
    "教養",
    "資格",
    "実験",
    "研究",
  ];

  const NO_IMAGE =
    "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='640' height='480'%3E%3Crect width='100%25' height='100%25' fill='%23f1f3f6'/%3E%3Ctext x='50%25' y='50%25' dominant-baseline='middle' text-anchor='middle' fill='%23707b88' font-size='26'%3ENo Image%3C/text%3E%3C/svg%3E";

  const API_TIMEOUT_MS = 45000;
  const API_RETRIES = 2;

  const BOOK_CACHE_PREFIX = "book-meta-cache-v5";
  const BOOK_CACHE_TTL_MS = 1000 * 60 * 60 * 24 * 180;
  const BOOK_TIMEOUT_STEPS = [22000, 36000, 52000, 70000];
  const BOOK_RETRY_COUNT = 2;
  const BOOK_RETRY_DELAY_MS = 1200;
  const BOOK_SUMMARY_FALLBACK = "概要は取得できませんでした。説明欄の追記をご利用ください。";
  const BOOK_MIN_SUMMARY_LENGTH = 24;
  const LISTINGS_CACHE_TTL_MS = 1000 * 60 * 8;
  const MAX_UPLOAD_IMAGES = 5;
  const MAX_UPLOAD_FILE_BYTES = 8 * 1024 * 1024;
  const IMAGE_UPLOAD_CONCURRENCY = 3;
  const UPLOAD_MAX_EDGE = 1800;
  const UPLOAD_JPEG_QUALITY = 0.84;

  const state = {
    apiBaseUrl: sanitizeUrl(APP_CONFIG.APPS_SCRIPT_BASE_URL || FIXED_APPS_SCRIPT_BASE_URL),
    session: loadSession(),
    activeTab: loadActiveTab(),
    filters: {
      category: "",
      q: "",
    },
    pendingResident: null,
    detailListingId: "",
    listings: [],
    myPage: {
      wanted: [],
      mine: [],
    },
    requests: {
      listings: null,
      myPage: null,
    },
    seller: {
      currentBookMeta: null,
      currentJan: "",
    },
  };

  const dom = {
    sessionBadge: byId("sessionBadge"),
    openLoginButton: byId("openLoginButton"),
    logoutButton: byId("logoutButton"),

    searchInput: byId("searchInput"),
    quickCategoryRow: byId("quickCategoryRow"),

    homeScreen: byId("homeScreen"),
    plusScreen: byId("plusScreen"),
    myScreen: byId("myScreen"),

    reloadButton: byId("reloadButton"),
    listingGrid: byId("listingGrid"),
    homeStatus: byId("homeStatus"),

    plusLocked: byId("plusLocked"),
    listingForm: byId("listingForm"),
    itemTypeInput: byId("itemTypeInput"),
    categoryInput: byId("categoryInput"),
    titleInput: byId("titleInput"),
    descriptionInput: byId("descriptionInput"),
    bookFields: byId("bookFields"),
    janInput: byId("janInput"),
    lookupBookButton: byId("lookupBookButton"),
    lookupStatus: byId("lookupStatus"),
    authorInput: byId("authorInput"),
    publisherInput: byId("publisherInput"),
    publishedDateInput: byId("publishedDateInput"),
    subjectTags: byId("subjectTags"),
    subjectFreeInput: byId("subjectFreeInput"),
    imageInput: byId("imageInput"),
    imagePreview: byId("imagePreview"),
    submitListingButton: byId("submitListingButton"),
    listingStatus: byId("listingStatus"),

    reloadMyPageButton: byId("reloadMyPageButton"),
    myLocked: byId("myLocked"),
    myStatus: byId("myStatus"),
    myWantedList: byId("myWantedList"),
    myWantedEmpty: byId("myWantedEmpty"),
    myListingsList: byId("myListingsList"),
    myListingsEmpty: byId("myListingsEmpty"),
    myConflictSection: byId("myConflictSection"),
    myConflictList: byId("myConflictList"),
    myConflictEmpty: byId("myConflictEmpty"),

    bottomNav: byId("bottomNav"),

    loginModal: byId("loginModal"),
    closeLoginButton: byId("closeLoginButton"),
    roomNumberInput: byId("roomNumberInput"),
    checkRoomButton: byId("checkRoomButton"),
    confirmLoginButton: byId("confirmLoginButton"),
    nameConfirmBox: byId("nameConfirmBox"),
    confirmName: byId("confirmName"),
    loginStatus: byId("loginStatus"),

    wishersModal: byId("wishersModal"),
    closeWishersButton: byId("closeWishersButton"),
    wishersTitle: byId("wishersTitle"),
    wishersList: byId("wishersList"),
    wishersStatus: byId("wishersStatus"),

    detailModal: byId("detailModal"),
    closeDetailButton: byId("closeDetailButton"),
    detailImage: byId("detailImage"),
    detailTitle: byId("detailTitle"),
    detailMeta: byId("detailMeta"),
    detailTags: byId("detailTags"),
    detailDescription: byId("detailDescription"),
    detailStatus: byId("detailStatus"),
    detailWantButton: byId("detailWantButton"),
    detailWishersButton: byId("detailWishersButton"),
  };

  window.__dormMarketDebug = state;

  try {
    initialize();
  } catch (error) {
    console.error("Initialization failed", error);
    alert("初期化に失敗しました。ページを再読み込みしてください。\n" + String(error.message || error));
  }

  function initialize() {
    ensureRequiredDom();
    renderSubjectTags();
    bindEvents();
    renderSessionState();
    switchTab(state.activeTab, { persist: false });
    hydrateListingsFromCache();
    refreshListings();
  }

  function ensureRequiredDom() {
    const required = [
      "openLoginButton",
      "searchInput",
      "quickCategoryRow",
      "bottomNav",
      "homeScreen",
      "plusScreen",
      "myScreen",
      "listingGrid",
      "listingForm",
      "myConflictSection",
      "loginModal",
      "wishersModal",
      "detailModal",
    ];

    const missing = required.filter((key) => !dom[key]);
    if (missing.length) {
      throw new Error("DOM element not found: " + missing.join(", "));
    }
  }

  function bindEvents() {
    dom.openLoginButton.addEventListener("click", openLoginModal);
    dom.closeLoginButton.addEventListener("click", closeLoginModal);
    dom.logoutButton.addEventListener("click", handleLogout);

    dom.loginModal.addEventListener("click", (event) => {
      if (event.target === dom.loginModal) {
        closeLoginModal();
      }
    });

    dom.checkRoomButton.addEventListener("click", handleCheckRoom);
    dom.confirmLoginButton.addEventListener("click", handleConfirmLogin);

    dom.roomNumberInput.addEventListener("input", () => {
      dom.roomNumberInput.value = normalizeRoomNumber(dom.roomNumberInput.value);
    });

    dom.roomNumberInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        handleCheckRoom();
      }
    });

    let searchTimer = null;
    dom.searchInput.addEventListener("input", () => {
      if (searchTimer) {
        window.clearTimeout(searchTimer);
      }
      searchTimer = window.setTimeout(() => {
        state.filters.q = normalizeText(dom.searchInput.value);
        renderHome();
      }, 220);
    });

    dom.quickCategoryRow.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-category]");
      if (!button) {
        return;
      }
      state.filters.category = normalizeText(button.dataset.category);
      applyQuickCategoryButtons();
      renderHome();
    });

    dom.reloadButton.addEventListener("click", refreshListings);
    dom.listingGrid.addEventListener("click", handleListingGridClick);
    dom.listingGrid.addEventListener("keydown", handleListingGridKeyDown);
    dom.myWantedList.addEventListener("click", handleListingActionClick);
    dom.myListingsList.addEventListener("click", handleListingActionClick);
    dom.myConflictList.addEventListener("click", handleListingActionClick);

    dom.itemTypeInput.addEventListener("change", syncBookFieldVisibility);
    dom.lookupBookButton.addEventListener("click", handleBookLookup);
    dom.janInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        handleBookLookup();
      }
    });

    dom.imageInput.addEventListener("change", handleImagePreview);
    dom.listingForm.addEventListener("submit", handleSubmitListing);

    dom.reloadMyPageButton.addEventListener("click", () => {
      if (!state.session) {
        openLoginModal();
        return;
      }
      refreshMyPage();
    });

    dom.bottomNav.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-tab]");
      if (!button) {
        return;
      }
      const tab = button.dataset.tab;
      if (tab === "my" && !state.session) {
        switchTab("my");
        openLoginModal();
        return;
      }
      if (tab === "plus" && !state.session) {
        switchTab("plus");
        openLoginModal();
        return;
      }
      switchTab(tab);
    });

    dom.closeWishersButton.addEventListener("click", closeWishersModal);
    dom.wishersModal.addEventListener("click", (event) => {
      if (event.target === dom.wishersModal) {
        closeWishersModal();
      }
    });

    dom.closeDetailButton.addEventListener("click", closeDetailModal);
    dom.detailModal.addEventListener("click", (event) => {
      if (event.target === dom.detailModal) {
        closeDetailModal();
      }
    });
    dom.detailWantButton.addEventListener("click", handleListingActionClick);
    dom.detailWishersButton.addEventListener("click", handleListingActionClick);
    document.addEventListener("keydown", handleGlobalKeyDown);
  }

  function byId(id) {
    return document.getElementById(id);
  }

  function handleGlobalKeyDown(event) {
    if (event.key !== "Escape") {
      return;
    }
    if (!dom.detailModal.classList.contains("hidden")) {
      closeDetailModal();
      return;
    }
    if (!dom.wishersModal.classList.contains("hidden")) {
      closeWishersModal();
      return;
    }
    if (!dom.loginModal.classList.contains("hidden")) {
      closeLoginModal();
    }
  }

  function switchTab(tab, options = {}) {
    const valid = ["home", "plus", "my"];
    const target = valid.includes(tab) ? tab : "home";
    state.activeTab = target;

    if (options.persist !== false) {
      saveActiveTab(target);
    }

    dom.homeScreen.classList.toggle("is-active", target === "home");
    dom.homeScreen.classList.toggle("hidden", target !== "home");
    dom.plusScreen.classList.toggle("is-active", target === "plus");
    dom.plusScreen.classList.toggle("hidden", target !== "plus");
    dom.myScreen.classList.toggle("is-active", target === "my");
    dom.myScreen.classList.toggle("hidden", target !== "my");

    Array.from(dom.bottomNav.querySelectorAll(".nav-item")).forEach((button) => {
      const isActive = button.dataset.tab === target;
      button.classList.toggle("is-active", isActive);
    });

    dom.quickCategoryRow.classList.toggle("hidden", target !== "home");

    applyAccessState();

    if (target === "my") {
      refreshMyPage();
    }
  }

  function renderSessionState() {
    if (!state.session) {
      dom.sessionBadge.textContent = "未ログイン";
      dom.openLoginButton.classList.remove("hidden");
      dom.logoutButton.classList.add("hidden");
      applyAccessState();
      syncDetailModal();
      return;
    }

    dom.sessionBadge.textContent = `${state.session.roomNumber}号室 ${state.session.name}`;
    dom.openLoginButton.classList.add("hidden");
    dom.logoutButton.classList.remove("hidden");
    applyAccessState();
    syncDetailModal();
  }

  function canCreateListing(session = state.session) {
    return Boolean(session && session.userId);
  }

  function canViewConflictSection(session = state.session) {
    return Boolean(session && session.userId);
  }

  function applyAccessState() {
    const loggedIn = Boolean(state.session);
    const canPost = canCreateListing();

    dom.myLocked.classList.toggle("hidden", loggedIn);

    dom.plusLocked.classList.toggle("hidden", canPost);
    if (!loggedIn) {
      dom.plusLocked.textContent = "出品するにはログインしてください。";
    } else {
      dom.plusLocked.textContent = "";
    }

    dom.submitListingButton.disabled = !canPost;
    dom.myConflictSection.classList.toggle("hidden", !canViewConflictSection());
  }

  function applyQuickCategoryButtons() {
    Array.from(dom.quickCategoryRow.querySelectorAll("button[data-category]")).forEach((button) => {
      const active = normalizeText(button.dataset.category) === state.filters.category;
      button.classList.toggle("is-active", active);
    });
  }

  function hydrateListingsFromCache() {
    const cached = loadListingsCache();
    if (!cached.length) {
      return;
    }
    state.listings = cached;
    renderHome();
    setStatus(dom.homeStatus, "前回データを表示中...");
  }

  function loadListingsCache() {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEYS.listingsCache);
      if (!raw) {
        return [];
      }
      const parsed = JSON.parse(raw);
      const savedAt = Number(parsed?.savedAt || 0);
      const rows = Array.isArray(parsed?.listings) ? parsed.listings : [];
      if (!savedAt || Date.now() - savedAt > LISTINGS_CACHE_TTL_MS) {
        window.localStorage.removeItem(STORAGE_KEYS.listingsCache);
        return [];
      }
      return rows.map(normalizeListing).filter((item) => item.listingId);
    } catch {
      return [];
    }
  }

  function writeListingsCache(listings) {
    try {
      window.localStorage.setItem(
        STORAGE_KEYS.listingsCache,
        JSON.stringify({
          savedAt: Date.now(),
          listings: Array.isArray(listings) ? listings : [],
        })
      );
    } catch {
      // ignore quota errors
    }
  }

  function openLoginModal() {
    state.pendingResident = null;
    dom.roomNumberInput.value = "";
    dom.nameConfirmBox.classList.add("hidden");
    dom.confirmLoginButton.classList.add("hidden");
    setStatus(dom.loginStatus, "");
    dom.loginModal.classList.remove("hidden");
    dom.roomNumberInput.focus();
  }

  function closeLoginModal() {
    dom.loginModal.classList.add("hidden");
  }

  async function handleCheckRoom() {
    const roomNumber = normalizeRoomNumber(dom.roomNumberInput.value);
    dom.roomNumberInput.value = roomNumber;

    if (roomNumber.length !== 3) {
      setStatus(dom.loginStatus, "部屋番号は3桁で入力してください。", "error");
      return;
    }

    dom.checkRoomButton.disabled = true;
    dom.confirmLoginButton.disabled = true;
    setStatus(dom.loginStatus, "照合中です...");

    try {
      const result = await apiPost("loginByRoom", { roomNumber });
      if (!result.ok || !result.resident) {
        throw new Error(result.error || "部屋番号が見つかりません。");
      }

      state.pendingResident = result.resident;
      dom.confirmName.textContent = result.resident.name;
      dom.nameConfirmBox.classList.remove("hidden");
      dom.confirmLoginButton.classList.remove("hidden");
      setStatus(dom.loginStatus, "名前を確認しました。ログインしてください。", "ok");
    } catch (error) {
      state.pendingResident = null;
      dom.nameConfirmBox.classList.add("hidden");
      dom.confirmLoginButton.classList.add("hidden");
      setStatus(dom.loginStatus, String(error.message || error), "error");
    } finally {
      dom.checkRoomButton.disabled = false;
      dom.confirmLoginButton.disabled = false;
    }
  }

  function handleConfirmLogin() {
    if (!state.pendingResident) {
      setStatus(dom.loginStatus, "先に名前確認をしてください。", "error");
      return;
    }

    state.session = {
      userId: normalizeText(state.pendingResident.userId),
      roomNumber: normalizeText(state.pendingResident.roomNumber),
      name: normalizeText(state.pendingResident.name),
      role: "viewer",
      loggedInAt: new Date().toISOString(),
    };

    saveSession(state.session);
    renderSessionState();
    closeLoginModal();
    refreshListings();
  }

  function handleLogout() {
    state.session = null;
    clearSession();
    renderSessionState();
    renderMyPage([], [], []);
    closeWishersModal();
    closeDetailModal();
    setStatus(dom.myStatus, "");
    switchTab("home");
    refreshListings();
  }

  async function refreshListings() {
    if (state.requests.listings) {
      return state.requests.listings;
    }

    state.requests.listings = (async () => {
      setStatus(dom.homeStatus, "一覧を読み込み中...");
      const result = await apiGet("listListings", {
        viewerId: state.session ? state.session.userId : "",
      });

      if (!result.ok) {
        if (!state.listings.length) {
          dom.listingGrid.innerHTML = "";
          closeDetailModal();
        }
        setStatus(dom.homeStatus, result.error || "一覧の取得に失敗しました。", "error");
        return;
      }

      const rawListings = Array.isArray(result.listings)
        ? result.listings
        : Array.isArray(result.books)
          ? result.books
          : [];

      state.listings = rawListings.map(normalizeListing).filter((item) => item.listingId);
      writeListingsCache(rawListings);
      renderHome();

      if (state.activeTab === "my" && state.session) {
        await refreshMyPage();
      }
    })();

    try {
      return await state.requests.listings;
    } finally {
      state.requests.listings = null;
    }
  }

  function renderHome() {
    const filtered = filterListings(state.listings)
      .slice()
      .sort((a, b) => toTimestamp(b.createdAt) - toTimestamp(a.createdAt));

    dom.listingGrid.innerHTML = "";

    if (!filtered.length) {
      syncDetailModal();
      setStatus(dom.homeStatus, "0件です。", "ok");
      return;
    }

    const fragment = document.createDocumentFragment();
    filtered.forEach((listing) => {
      fragment.appendChild(createListingCard(listing));
    });
    dom.listingGrid.appendChild(fragment);

    syncDetailModal();
    setStatus(dom.homeStatus, `${filtered.length}件を表示中`, "ok");
  }

  function filterListings(listings) {
    const category = state.filters.category;
    const queryTokens = parseSearchTokens(state.filters.q);

    return listings.filter((listing) => {
      if (category && normalizeItemType(listing.itemType) !== category) {
        return false;
      }

      if (queryTokens.length) {
        const rawIndex = normalizeSearchText(listing.searchRaw || buildListingSearchText(listing));
        const phoneticIndex = listing.searchPhonetic || toPhoneticSearchKey(rawIndex);

        const matched = queryTokens.every((token) => {
          const rawMatch = token.raw ? rawIndex.includes(token.raw) : false;
          const phoneticMatch = token.phonetic ? phoneticIndex.includes(token.phonetic) : false;
          return rawMatch || phoneticMatch;
        });

        if (!matched) {
          return false;
        }
      }

      return true;
    });
  }

  function createListingCard(listing) {
    const card = document.createElement("article");
    card.className = "listing-card";
    card.dataset.listingId = listing.listingId;
    card.tabIndex = 0;
    card.setAttribute("role", "button");
    card.setAttribute("aria-label", `${listing.title || "商品"} の詳細を開く`);

    const photo = document.createElement("img");
    photo.className = "listing-photo";
    photo.alt = `${listing.title || "商品"} の画像`;
    photo.loading = "lazy";
    setImageWithFallback(photo, imageCandidatesFromListing(listing));

    const title = document.createElement("h3");
    title.className = "listing-title";
    title.textContent = listing.title || "タイトル未設定";

    const meta = document.createElement("p");
    meta.className = "muted";
    meta.textContent = [
      listing.category,
      `欲しい ${Number(listing.wantedCount || 0)}名`,
      listing.author,
      listing.publisher,
      listing.sellerName ? `出品: ${listing.sellerName}` : "",
    ]
      .filter(Boolean)
      .join(" / ");

    const tags = document.createElement("div");
    tags.className = "tag-row";

    appendTag(tags, listing.itemTypeLabel || listingTypeLabel(listing.itemType));
    if (listing.jan) {
      appendTag(tags, "JANあり");
    }
    listing.subjectTags.forEach((subject) => appendTag(tags, subject, "subject"));

    card.appendChild(photo);
    card.appendChild(title);
    if (meta.textContent) {
      card.appendChild(meta);
    }
    card.appendChild(tags);

    return card;
  }

  function appendTag(container, text, className = "") {
    if (!normalizeText(text)) {
      return;
    }
    const span = document.createElement("span");
    span.className = `tag ${className}`.trim();
    span.textContent = text;
    container.appendChild(span);
  }

  function listingTypeLabel(itemType) {
    if (itemType === "book") return "本";
    if (itemType === "goods") return "小物";
    if (itemType === "bicycle") return "自転車";
    return "その他";
  }

  function handleListingGridClick(event) {
    const card = event.target.closest(".listing-card[data-listing-id]");
    if (!card) {
      return;
    }
    openDetailByListingId(card.dataset.listingId);
  }

  function handleListingGridKeyDown(event) {
    if (event.key !== "Enter" && event.key !== " ") {
      return;
    }
    const card = event.target.closest(".listing-card[data-listing-id]");
    if (!card) {
      return;
    }
    event.preventDefault();
    openDetailByListingId(card.dataset.listingId);
  }

  function openDetailByListingId(listingId) {
    const listing = findListingById(normalizeText(listingId));
    if (!listing) {
      return;
    }
    state.detailListingId = listing.listingId;
    renderDetailModal(listing);
    dom.detailModal.classList.remove("hidden");
  }

  function closeDetailModal() {
    state.detailListingId = "";
    dom.detailModal.classList.add("hidden");
    dom.detailTitle.textContent = "-";
    dom.detailMeta.textContent = "-";
    dom.detailDescription.textContent = "-";
    dom.detailImage.onerror = null;
    dom.detailImage.src = NO_IMAGE;
    dom.detailTags.innerHTML = "";
    dom.detailWantButton.dataset.listingId = "";
    dom.detailWishersButton.dataset.listingId = "";
    setStatus(dom.detailStatus, "");
  }

  function syncDetailModal() {
    if (!state.detailListingId) {
      return;
    }
    const listing = findListingById(state.detailListingId);
    if (!listing) {
      closeDetailModal();
      return;
    }
    renderDetailModal(listing);
  }

  function renderDetailModal(listing) {
    dom.detailTitle.textContent = listing.title || "タイトル未設定";
    dom.detailMeta.textContent = [
      listing.category,
      listing.sellerName ? `出品: ${listing.sellerName}` : "",
      listing.author ? `著者: ${listing.author}` : "",
      listing.publisher ? `出版社: ${listing.publisher}` : "",
      listing.publishedDate ? `出版: ${listing.publishedDate}` : "",
      `欲しい ${Number(listing.wantedCount || 0)}名`,
    ]
      .filter(Boolean)
      .join(" / ");
    dom.detailDescription.textContent = listing.description || listing.summary || "説明は未入力です。";

    dom.detailImage.alt = `${listing.title || "商品"} の画像`;
    setImageWithFallback(dom.detailImage, imageCandidatesFromListing(listing));

    dom.detailTags.innerHTML = "";
    appendTag(dom.detailTags, listing.itemTypeLabel || listingTypeLabel(listing.itemType));
    if (listing.jan) {
      appendTag(dom.detailTags, "JANあり");
    }
    listing.subjectTags.forEach((subject) => appendTag(dom.detailTags, subject, "subject"));

    dom.detailWantButton.dataset.listingId = listing.listingId;
    dom.detailWishersButton.dataset.listingId = listing.listingId;
    dom.detailWantButton.textContent = listing.alreadyWanted
      ? "欲しい解除"
      : state.session
        ? "欲しい"
        : "ログインして欲しい";
    dom.detailWishersButton.textContent = `欲しい人 ${Number(listing.wantedCount || 0)}名`;

    setStatus(dom.detailStatus, "");
  }

  async function handleListingActionClick(event) {
    const button = event.target.closest("button[data-action]");
    if (!button) {
      return;
    }

    const action = button.dataset.action;
    const listingId = normalizeText(button.dataset.listingId);
    const listing = findListingById(listingId);

    if (!listing) {
      return;
    }

    if (action === "toggle-want") {
      await toggleWant(listing);
      return;
    }

    if (action === "remove-want") {
      await removeWant(listing);
      return;
    }

    if (action === "show-wishers") {
      openWishersModal(listing);
    }
  }

  function findListingById(listingId) {
    const inMain = state.listings.find((item) => item.listingId === listingId);
    if (inMain) {
      return inMain;
    }

    const inWanted = state.myPage.wanted.find((item) => item.listingId === listingId);
    if (inWanted) {
      return inWanted;
    }

    const inMine = state.myPage.mine.find((item) => item.listingId === listingId);
    return inMine || null;
  }

  async function toggleWant(listing) {
    if (!state.session) {
      openLoginModal();
      return;
    }

    setStatus(dom.homeStatus, "更新中...");
    if (state.detailListingId === listing.listingId) {
      setStatus(dom.detailStatus, "更新中...");
    }

    let result;
    if (listing.alreadyWanted) {
      result = await apiPost("removeInterest", {
        viewerId: state.session.userId,
        listingId: listing.listingId,
      });
    } else {
      result = await apiPost("addInterest", {
        viewerId: state.session.userId,
        viewerName: state.session.name,
        listingId: listing.listingId,
      });
    }

    if (!result.ok) {
      setStatus(dom.homeStatus, result.error || "更新に失敗しました。", "error");
      if (state.detailListingId === listing.listingId) {
        setStatus(dom.detailStatus, result.error || "更新に失敗しました。", "error");
      }
      return;
    }

    await refreshListings();
    if (state.detailListingId === listing.listingId) {
      setStatus(dom.detailStatus, "更新しました。", "ok");
    }
  }

  async function removeWant(listing) {
    if (!state.session) {
      openLoginModal();
      return;
    }

    const result = await apiPost("removeInterest", {
      viewerId: state.session.userId,
      listingId: listing.listingId,
    });

    if (!result.ok) {
      setStatus(dom.myStatus, result.error || "欲しい解除に失敗しました。", "error");
      return;
    }

    await refreshListings();
  }

  function openWishersModal(listing) {
    dom.wishersTitle.textContent = listing.title || "商品";
    dom.wishersList.innerHTML = "";

    const users = Array.isArray(listing.wantedBy) ? listing.wantedBy : [];
    if (!users.length) {
      const li = document.createElement("li");
      li.textContent = "まだ希望者はいません。";
      dom.wishersList.appendChild(li);
      setStatus(dom.wishersStatus, "");
    } else if (!state.session) {
      const li = document.createElement("li");
      li.textContent = `希望者は ${users.length}名です。ログインすると名前を確認できます。`;
      dom.wishersList.appendChild(li);
      setStatus(dom.wishersStatus, "");
    } else {
      users.forEach((user) => {
        const li = document.createElement("li");
        li.textContent = user.viewerName || user.viewerId;
        dom.wishersList.appendChild(li);
      });
      setStatus(dom.wishersStatus, `${users.length}名が希望しています。`, "ok");
    }

    dom.wishersModal.classList.remove("hidden");
  }

  function closeWishersModal() {
    dom.wishersModal.classList.add("hidden");
    dom.wishersTitle.textContent = "-";
    dom.wishersList.innerHTML = "";
    setStatus(dom.wishersStatus, "");
  }

  function syncBookFieldVisibility() {
    const isBook = dom.itemTypeInput.value === "book";
    dom.bookFields.classList.toggle("hidden", !isBook);
    if (!isBook) {
      dom.janInput.value = "";
      state.seller.currentBookMeta = null;
      state.seller.currentJan = "";
      setStatus(dom.lookupStatus, "");
    }
  }

  function renderSubjectTags() {
    dom.subjectTags.innerHTML = "";
    SUBJECT_OPTIONS.forEach((subject) => {
      const label = document.createElement("label");
      label.className = "subject-chip";

      const input = document.createElement("input");
      input.type = "checkbox";
      input.value = subject;

      const text = document.createElement("span");
      text.textContent = subject;

      label.appendChild(input);
      label.appendChild(text);
      dom.subjectTags.appendChild(label);
    });

    syncBookFieldVisibility();
  }

  async function handleBookLookup() {
    const jan = normalizeJan(dom.janInput.value);
    dom.janInput.value = jan;

    if (!isJanBookCode(jan)) {
      setStatus(dom.lookupStatus, "JAN/ISBN-13 を正しく入力してください。", "error");
      return;
    }

    dom.lookupBookButton.disabled = true;
    setStatus(dom.lookupStatus, "書誌情報を取得中...");

    try {
      const meta = await lookupBookByJan(jan, { forceRefresh: false });
      state.seller.currentBookMeta = meta;
      state.seller.currentJan = jan;

      if (!normalizeText(dom.titleInput.value) && meta.title) {
        dom.titleInput.value = meta.title;
      }
      if (!normalizeText(dom.descriptionInput.value) && meta.summary) {
        dom.descriptionInput.value = meta.summary;
      }
      if (!normalizeText(dom.authorInput.value) && meta.author) {
        dom.authorInput.value = meta.author;
      }
      if (!normalizeText(dom.publisherInput.value) && meta.publisher) {
        dom.publisherInput.value = meta.publisher;
      }
      if (!normalizeText(dom.publishedDateInput.value) && meta.publishedDate) {
        dom.publishedDateInput.value = meta.publishedDate;
      }

      const hasCover = Boolean(meta.coverUrl || meta.coverCandidates.length);
      const hasSummary = meta.summary && meta.summary.length > 0;
      if (hasCover && hasSummary) {
        setStatus(dom.lookupStatus, "書誌情報を取得しました。", "ok");
      } else {
        setStatus(dom.lookupStatus, "一部の情報のみ取得しました。必要なら写真を追加してください。", "ok");
      }
    } catch (error) {
      setStatus(dom.lookupStatus, String(error.message || error), "error");
    } finally {
      dom.lookupBookButton.disabled = false;
    }
  }

  function handleImagePreview() {
    dom.imagePreview.innerHTML = "";
    const files = Array.from(dom.imageInput.files || []);

    if (files.length > MAX_UPLOAD_IMAGES) {
      setStatus(dom.listingStatus, "画像は最大5枚までです。", "error");
      dom.imageInput.value = "";
      return;
    }

    files.forEach((file) => {
      const img = document.createElement("img");
      img.alt = file.name;
      const reader = new FileReader();
      reader.onload = () => {
        img.src = String(reader.result || NO_IMAGE);
      };
      reader.readAsDataURL(file);
      dom.imagePreview.appendChild(img);
    });

    if (files.length) {
      setStatus(dom.listingStatus, `${files.length}枚選択中`);
    } else {
      setStatus(dom.listingStatus, "");
    }
  }

  async function handleSubmitListing(event) {
    event.preventDefault();

    if (!state.session) {
      openLoginModal();
      return;
    }
    if (!canCreateListing()) {
      setStatus(dom.listingStatus, "出品するにはログインしてください。", "error");
      return;
    }

    const itemType = normalizeItemType(dom.itemTypeInput.value);
    const category = normalizeText(dom.categoryInput.value) || defaultCategory(itemType);
    const title = normalizeText(dom.titleInput.value);
    const description = normalizeText(dom.descriptionInput.value);

    if (!category) {
      setStatus(dom.listingStatus, "カテゴリを入力してください。", "error");
      return;
    }
    if (!title) {
      setStatus(dom.listingStatus, "タイトルを入力してください。", "error");
      return;
    }

    dom.submitListingButton.disabled = true;

    try {
      let jan = "";
      let author = "";
      let publisher = "";
      let publishedDate = "";
      let subjectTags = [];
      let fallbackCoverUrl = "";

      if (itemType === "book") {
        jan = normalizeJan(dom.janInput.value);
        if (jan && !isJanBookCode(jan)) {
          throw new Error("JAN/ISBN-13 が正しくありません。");
        }

        const activeBookMeta =
          jan && state.seller.currentBookMeta && state.seller.currentJan === jan ? state.seller.currentBookMeta : null;
        if (activeBookMeta) {
          fallbackCoverUrl = normalizeText(activeBookMeta.coverUrl || activeBookMeta.coverCandidates?.[0]);
        }

        author = normalizeText(dom.authorInput.value) || normalizeText(activeBookMeta?.author);
        publisher = normalizeText(dom.publisherInput.value) || normalizeText(activeBookMeta?.publisher);
        publishedDate =
          normalizeText(dom.publishedDateInput.value) || normalizeText(activeBookMeta?.publishedDate);
        subjectTags = collectSubjectTags();
      }

      const payloads = await buildUploadPayloads(dom.imageInput.files);
      const imageUrls = await uploadImagePayloads(payloads);

      if (!imageUrls.length && fallbackCoverUrl) {
        imageUrls.push(fallbackCoverUrl);
      }

      const listing = {
        itemType,
        category,
        title,
        description,
        imageUrls,
        jan,
        subjectTags,
        author,
        publisher,
        publishedDate,
      };

      setStatus(dom.listingStatus, "出品を登録中...");
      const result = await apiPost(
        "addListing",
        {
          sellerId: state.session.userId,
          sellerName: state.session.name,
          listing,
        },
        60000,
        1
      );

      if (!result.ok) {
        throw new Error(result.error || "出品登録に失敗しました。");
      }

      dom.listingForm.reset();
      dom.imagePreview.innerHTML = "";
      state.seller.currentBookMeta = null;
      state.seller.currentJan = "";
      syncBookFieldVisibility();
      setStatus(dom.lookupStatus, "");
      setStatus(dom.listingStatus, "出品しました。", "ok");

      switchTab("home");
      await refreshListings();
    } catch (error) {
      setStatus(dom.listingStatus, String(error.message || error), "error");
    } finally {
      applyAccessState();
    }
  }

  function collectSubjectTags() {
    const checked = Array.from(dom.subjectTags.querySelectorAll("input[type='checkbox']:checked")).map((input) =>
      normalizeText(input.value)
    );
    const free = normalizeText(dom.subjectFreeInput.value)
      .split(",")
      .map((token) => normalizeText(token))
      .filter(Boolean);

    return uniqueStrings([...checked, ...free]);
  }

  async function refreshMyPage() {
    if (!state.session) {
      renderMyPage([], [], []);
      return;
    }

    if (state.requests.myPage) {
      return state.requests.myPage;
    }

    state.requests.myPage = (async () => {
      setStatus(dom.myStatus, "マイページを読み込み中...");
      const result = await apiGet("listMyPage", { viewerId: state.session.userId }, 45000, 2);

      if (!result.ok) {
        setStatus(dom.myStatus, result.error || "マイページ取得に失敗しました。", "error");
        return;
      }

      state.myPage.wanted = (result.wantedListings || []).map(normalizeListing).filter((item) => item.listingId);
      state.myPage.mine = (result.myListings || []).map(normalizeListing).filter((item) => item.listingId);

      renderMyPage(state.myPage.wanted, state.myPage.mine, state.myPage.mine);
    })();

    try {
      return await state.requests.myPage;
    } finally {
      state.requests.myPage = null;
    }
  }

  function renderMyPage(wantedListings, myListings, conflictSourceListings = []) {
    dom.myWantedList.innerHTML = "";
    dom.myListingsList.innerHTML = "";
    dom.myConflictList.innerHTML = "";

    if (!state.session) {
      dom.myLocked.classList.remove("hidden");
      dom.myWantedEmpty.classList.remove("hidden");
      dom.myListingsEmpty.classList.remove("hidden");
      dom.myConflictEmpty.classList.add("hidden");
      applyAccessState();
      setStatus(dom.myStatus, "ログインしてください。", "error");
      return;
    }

    dom.myLocked.classList.add("hidden");
    applyAccessState();

    if (!wantedListings.length) {
      dom.myWantedEmpty.classList.remove("hidden");
    } else {
      dom.myWantedEmpty.classList.add("hidden");
      wantedListings.forEach((listing) => {
        dom.myWantedList.appendChild(createStackItem(listing, [
          {
            action: "remove-want",
            label: "欲しい解除",
            className: "action-btn",
          },
          {
            action: "show-wishers",
            label: "欲しい人を見る",
            className: "action-btn",
          },
        ]));
      });
    }

    if (!myListings.length) {
      dom.myListingsEmpty.classList.remove("hidden");
    } else {
      dom.myListingsEmpty.classList.add("hidden");
      myListings.forEach((listing) => {
        dom.myListingsList.appendChild(createStackItem(listing, [
          {
            action: "show-wishers",
            label: `欲しい人 ${Number(listing.wantedCount || 0)}名`,
            className: "action-btn",
          },
        ]));
      });
    }

    const conflicts = (Array.isArray(conflictSourceListings) ? conflictSourceListings : [])
      .filter((listing) => Number(listing.wantedCount || 0) > 0)
      .sort((a, b) => Number(b.wantedCount || 0) - Number(a.wantedCount || 0));

    if (canViewConflictSection()) {
      if (!conflicts.length) {
        dom.myConflictEmpty.classList.remove("hidden");
      } else {
        dom.myConflictEmpty.classList.add("hidden");
        conflicts.forEach((listing) => {
          dom.myConflictList.appendChild(createStackItem(listing, [
            {
              action: "show-wishers",
              label: "希望者を見る",
              className: "action-btn",
            },
          ]));
        });
      }
    } else {
      dom.myConflictEmpty.classList.add("hidden");
    }

    setStatus(dom.myStatus, "マイページを更新しました。", "ok");
  }

  function createStackItem(listing, actions) {
    const item = document.createElement("article");
    item.className = "stack-item";

    const head = document.createElement("div");
    head.className = "stack-head";

    const title = document.createElement("h4");
    title.textContent = listing.title || "タイトル未設定";

    const meta = document.createElement("span");
    meta.className = "tag";
    meta.textContent = `${listing.category || "カテゴリなし"} / 欲しい ${Number(listing.wantedCount || 0)}名`;

    head.appendChild(title);
    head.appendChild(meta);

    const desc = document.createElement("p");
    desc.className = "muted";
    desc.textContent = listing.description || listing.summary || "説明なし";

    const actionRow = document.createElement("div");
    actionRow.className = "card-actions";

    actions.forEach((actionItem) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = actionItem.className || "action-btn";
      button.dataset.action = actionItem.action;
      button.dataset.listingId = listing.listingId;
      button.textContent = actionItem.label;
      actionRow.appendChild(button);
    });

    item.appendChild(head);
    item.appendChild(desc);
    item.appendChild(actionRow);
    return item;
  }

  async function apiGet(action, params = {}, timeoutMs = API_TIMEOUT_MS, retries = API_RETRIES) {
    const url = new URL(state.apiBaseUrl);
    url.searchParams.set("action", action);

    Object.entries(params).forEach(([key, value]) => {
      if (value === undefined || value === null || value === "") {
        return;
      }
      url.searchParams.set(key, String(value));
    });

    return requestJson(url.toString(), { method: "GET" }, timeoutMs, retries);
  }

  async function apiPost(action, payload = {}, timeoutMs = API_TIMEOUT_MS, retries = API_RETRIES) {
    const body = new URLSearchParams();
    body.set("action", action);
    body.set("payload", JSON.stringify(payload || {}));

    return requestJson(
      state.apiBaseUrl,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        },
        body: body.toString(),
      },
      timeoutMs,
      retries
    );
  }

  async function requestJson(url, options, timeoutMs, retries) {
    for (let attempt = 0; attempt <= retries; attempt += 1) {
      try {
        const response = await fetchWithTimeout(url, options, timeoutMs);
        let json;
        try {
          json = await response.json();
        } catch {
          json = null;
        }

        if (!response.ok) {
          const message = json && json.error ? json.error : `HTTP ${response.status}`;
          const retryable = [408, 425, 429, 500, 502, 503, 504].includes(response.status);
          if (retryable && attempt < retries) {
            await wait(700 * (attempt + 1));
            continue;
          }
          return { ok: false, error: message };
        }

        return json || { ok: false, error: "invalid response" };
      } catch (error) {
        if (attempt >= retries) {
          return { ok: false, error: String(error.message || error) };
        }
        await wait(700 * (attempt + 1));
      }
    }

    return { ok: false, error: "unknown request error" };
  }

  async function fetchWithTimeout(url, options, timeoutMs) {
    const controller = new AbortController();
    const id = window.setTimeout(() => controller.abort(), timeoutMs);
    try {
      return await fetch(url, {
        ...options,
        signal: controller.signal,
      });
    } finally {
      window.clearTimeout(id);
    }
  }

  function normalizeListing(raw) {
    const listingId = normalizeText(raw?.listingId || raw?.listing_id || raw?.bookId || raw?.book_id);
    const itemType = normalizeItemType(raw?.itemType || raw?.item_type || "book");
    const wantedBy = parseArray(raw?.wantedBy || raw?.wanted_by || raw?.interestedUsers).map((item) => ({
      viewerId: normalizeText(item?.viewerId || item?.viewer_id),
      viewerName: normalizeText(item?.viewerName || item?.viewer_name || item?.name),
    }));

    const wantedCount = Number(raw?.wantedCount || raw?.wanted_count || wantedBy.length || 0);

    const listing = {
      listingId,
      itemType,
      itemTypeLabel: listingTypeLabel(itemType),
      category: normalizeText(raw?.category),
      title: normalizeText(raw?.title),
      description: normalizeText(raw?.description),
      summary: normalizeText(raw?.summary),
      imageUrls: parseArray(raw?.imageUrls || raw?.image_urls || raw?.image_urls_json).map(normalizeUrl).filter(Boolean),
      coverUrl: normalizeUrl(raw?.coverUrl || raw?.cover_url),
      jan: normalizeJan(raw?.jan),
      subjectTags: parseArray(raw?.subjectTags || raw?.subject_tags || raw?.subject_tags_json).map(normalizeText),
      author: normalizeText(raw?.author),
      publisher: normalizeText(raw?.publisher),
      publishedDate: normalizeText(raw?.publishedDate || raw?.published_date),
      sellerName: normalizeText(raw?.sellerName || raw?.seller_name),
      sellerId: normalizeText(raw?.sellerId || raw?.seller_id),
      status: normalizeText(raw?.status || "AVAILABLE"),
      createdAt: normalizeText(raw?.createdAt || raw?.created_at),
      wantedCount,
      wantedBy,
      alreadyWanted: Boolean(raw?.alreadyWanted || raw?.already_wanted),
    };

    listing.searchRaw = buildListingSearchText(listing);
    listing.searchPhonetic = toPhoneticSearchKey(listing.searchRaw);
    return listing;
  }

  function parseArray(value) {
    if (Array.isArray(value)) {
      return value;
    }
    if (typeof value === "string") {
      try {
        const parsed = JSON.parse(value);
        return Array.isArray(parsed) ? parsed : [];
      } catch {
        const text = normalizeText(value);
        if (!text) {
          return [];
        }
        if (/^(https?:\/\/|data:image\/)/i.test(text)) {
          return [text];
        }
        if (text.includes(",")) {
          return text
            .split(",")
            .map((token) => normalizeText(token))
            .filter(Boolean);
        }
        return [text];
      }
    }
    return [];
  }

  function buildListingSearchText(listing) {
    return normalizeSearchText(
      [
        listing.title,
        listing.description,
        listing.summary,
        listing.category,
        listing.itemTypeLabel || listingTypeLabel(listing.itemType),
        listing.sellerName,
        listing.author,
        listing.publisher,
        listing.jan,
        listing.subjectTags.join(" "),
      ].join(" ")
    );
  }

  function normalizeSearchText(value) {
    const text = normalizeText(value);
    if (!text) {
      return "";
    }
    return text.normalize("NFKC").toLowerCase().replace(/\s+/g, " ").trim();
  }

  function toPhoneticSearchKey(value) {
    const text = toHiragana(normalizeSearchText(value));
    if (!text) {
      return "";
    }

    const smallKanaMap = {
      ぁ: "あ",
      ぃ: "い",
      ぅ: "う",
      ぇ: "え",
      ぉ: "お",
      っ: "つ",
      ゃ: "や",
      ゅ: "ゆ",
      ょ: "よ",
      ゎ: "わ",
      ゕ: "か",
      ゖ: "け",
    };

    return text
      .normalize("NFD")
      .replace(/[\u3099\u309A]/g, "")
      .replace(/[ぁぃぅぇぉっゃゅょゎゕゖ]/g, (char) => smallKanaMap[char] || char)
      .replace(/[ーｰ〜～]/g, "")
      .replace(/[^0-9a-zぁ-ん]/g, "");
  }

  function parseSearchTokens(query) {
    const normalized = normalizeSearchText(query);
    if (!normalized) {
      return [];
    }
    const rawTokens = uniqueStrings(
      normalized
        .split(/\s+/)
        .map((token) => normalizeText(token))
        .filter(Boolean)
    );
    return rawTokens.map((raw) => ({
      raw,
      phonetic: toPhoneticSearchKey(raw),
    }));
  }

  function toHiragana(text) {
    return normalizeText(text).replace(/[ァ-ヶ]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 0x60));
  }

  function imageCandidatesFromListing(listing) {
    const originals = uniqueStrings([
      ...(Array.isArray(listing.imageUrls) ? listing.imageUrls : []),
      normalizeUrl(listing.coverUrl),
    ]);

    const expanded = [];
    originals.forEach((url) => {
      expandImageUrlCandidates(url).forEach((candidate) => {
        if (!candidate || expanded.includes(candidate)) {
          return;
        }
        expanded.push(candidate);
      });
    });
    return expanded.length ? expanded : [NO_IMAGE];
  }

  function expandImageUrlCandidates(url) {
    const normalized = normalizeUrl(url);
    if (!normalized) {
      return [];
    }

    const driveId = extractGoogleDriveFileId(normalized);
    if (!driveId) {
      return [normalized];
    }

    const encodedId = encodeURIComponent(driveId);
    return uniqueStrings([
      `https://lh3.googleusercontent.com/d/${encodedId}=w1600`,
      `https://drive.google.com/thumbnail?id=${encodedId}&sz=w1600`,
      `https://drive.google.com/uc?export=view&id=${encodedId}`,
      normalized,
    ]);
  }

  function extractGoogleDriveFileId(url) {
    const text = normalizeText(url);
    if (!text) {
      return "";
    }

    const patterns = [
      /[?&]id=([a-zA-Z0-9_-]{20,})/,
      /\/d\/([a-zA-Z0-9_-]{20,})/,
      /\/d\/([a-zA-Z0-9_-]{20,})=/,
    ];

    for (const pattern of patterns) {
      const match = text.match(pattern);
      if (match && match[1]) {
        return match[1];
      }
    }

    return "";
  }

  function setImageWithFallback(img, candidates) {
    const queue = Array.isArray(candidates) ? candidates.slice() : [];
    const next = () => {
      const candidate = queue.shift();
      if (!candidate) {
        img.onerror = null;
        img.src = NO_IMAGE;
        return;
      }
      img.onerror = next;
      img.src = candidate;
    };
    next();
  }

  function toTimestamp(value) {
    const date = new Date(value);
    const ms = date.getTime();
    return Number.isFinite(ms) ? ms : 0;
  }

  async function buildUploadPayloads(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) {
      return [];
    }
    if (files.length > MAX_UPLOAD_IMAGES) {
      throw new Error("画像は最大5枚までです。");
    }

    return await Promise.all(files.map((file, index) => buildUploadPayload(file, index)));
  }

  async function buildUploadPayload(file, index) {
    if (!file.type.startsWith("image/")) {
      throw new Error("画像ファイルのみアップロードできます。");
    }
    if (file.size > MAX_UPLOAD_FILE_BYTES) {
      throw new Error("画像は1枚8MB以下にしてください。");
    }

    const optimized = await optimizeImageForUpload(file);
    return {
      fileName: resolveUploadFileName(file.name, optimized.mimeType, index),
      mimeType: optimized.mimeType,
      dataUrl: optimized.dataUrl,
    };
  }

  async function uploadImagePayloads(payloads) {
    const list = Array.isArray(payloads) ? payloads : [];
    if (!list.length) {
      return [];
    }

    const queue = list.map((payload, index) => ({ payload, index }));
    const uploaded = new Array(list.length);
    let completed = 0;

    setStatus(dom.listingStatus, `画像アップロード中 (0/${list.length}) ...`);

    const workerCount = Math.min(IMAGE_UPLOAD_CONCURRENCY, queue.length);
    const workers = Array.from({ length: workerCount }, () =>
      (async () => {
        while (queue.length) {
          const task = queue.shift();
          if (!task) {
            return;
          }

          const upload = await apiPost("uploadImage", task.payload, 90000, 1);
          if (!upload.ok) {
            throw new Error(upload.error || "画像アップロードに失敗しました。");
          }

          uploaded[task.index] = normalizeUrl(upload.imageUrl);
          completed += 1;
          setStatus(dom.listingStatus, `画像アップロード中 (${completed}/${list.length}) ...`);
        }
      })()
    );

    await Promise.all(workers);
    return uploaded.filter(Boolean);
  }

  async function optimizeImageForUpload(file) {
    const originalMimeType = normalizeText(file.type) || "image/jpeg";
    const originalDataUrl = await readFileAsDataUrl(file);

    if (originalMimeType === "image/gif" || originalMimeType === "image/svg+xml") {
      return {
        mimeType: originalMimeType,
        dataUrl: originalDataUrl,
      };
    }

    let image;
    try {
      image = await loadImageElement(originalDataUrl);
    } catch {
      return {
        mimeType: originalMimeType,
        dataUrl: originalDataUrl,
      };
    }

    const width = Number(image.naturalWidth || image.width || 0);
    const height = Number(image.naturalHeight || image.height || 0);
    if (!width || !height) {
      return {
        mimeType: originalMimeType,
        dataUrl: originalDataUrl,
      };
    }

    const scale = Math.min(1, UPLOAD_MAX_EDGE / Math.max(width, height));
    if (scale >= 0.999 && file.size <= 1.5 * 1024 * 1024) {
      return {
        mimeType: originalMimeType,
        dataUrl: originalDataUrl,
      };
    }

    const targetWidth = Math.max(1, Math.round(width * scale));
    const targetHeight = Math.max(1, Math.round(height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = targetWidth;
    canvas.height = targetHeight;

    const ctx = canvas.getContext("2d");
    if (!ctx) {
      return {
        mimeType: originalMimeType,
        dataUrl: originalDataUrl,
      };
    }

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, targetWidth, targetHeight);
    ctx.drawImage(image, 0, 0, targetWidth, targetHeight);

    const compressedDataUrl = canvas.toDataURL("image/jpeg", UPLOAD_JPEG_QUALITY);
    if (!compressedDataUrl || compressedDataUrl.length >= originalDataUrl.length * 0.98) {
      return {
        mimeType: originalMimeType,
        dataUrl: originalDataUrl,
      };
    }

    return {
      mimeType: "image/jpeg",
      dataUrl: compressedDataUrl,
    };
  }

  function resolveUploadFileName(fileName, mimeType, index) {
    const rawName = normalizeText(fileName).replace(/\.[^.]+$/, "");
    const base = rawName || `image-${Date.now()}-${index + 1}`;
    const safe = base.replace(/[^\w.-]+/g, "_");
    if (mimeType === "image/png") {
      return safe.endsWith(".png") ? safe : `${safe}.png`;
    }
    if (mimeType === "image/webp") {
      return safe.endsWith(".webp") ? safe : `${safe}.webp`;
    }
    return safe.endsWith(".jpg") || safe.endsWith(".jpeg") ? safe : `${safe}.jpg`;
  }

  function loadImageElement(dataUrl) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error("画像の読み込みに失敗しました。"));
      img.src = dataUrl;
    });
  }

  function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(String(reader.result || ""));
      reader.onerror = () => reject(new Error("画像の読み込みに失敗しました。"));
      reader.readAsDataURL(file);
    });
  }

  function loadSession() {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEYS.session);
      if (!raw) {
        return null;
      }
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") {
        return null;
      }
      const userId = normalizeText(parsed.userId);
      const roomNumber = normalizeRoomNumber(parsed.roomNumber);
      const name = normalizeText(parsed.name);
      const role = normalizeRole(parsed.role);
      if (!userId || !roomNumber || !name) {
        return null;
      }
      return {
        userId,
        roomNumber,
        name,
        role,
        loggedInAt: normalizeText(parsed.loggedInAt),
      };
    } catch {
      return null;
    }
  }

  function saveSession(session) {
    window.localStorage.setItem(STORAGE_KEYS.session, JSON.stringify(session));
  }

  function clearSession() {
    window.localStorage.removeItem(STORAGE_KEYS.session);
  }

  function loadActiveTab() {
    const value = normalizeText(window.localStorage.getItem(STORAGE_KEYS.activeTab));
    if (["home", "plus", "my"].includes(value)) {
      return value;
    }
    return "home";
  }

  function saveActiveTab(tab) {
    window.localStorage.setItem(STORAGE_KEYS.activeTab, tab);
  }

  function setStatus(element, message = "", tone = "") {
    if (!element) {
      return;
    }

    element.textContent = message;
    element.classList.remove("is-error", "is-ok");
    if (tone === "error") {
      element.classList.add("is-error");
    }
    if (tone === "ok") {
      element.classList.add("is-ok");
    }
  }

  function normalizeText(value) {
    if (value === undefined || value === null) {
      return "";
    }
    return String(value).trim();
  }

  function sanitizeUrl(value) {
    return normalizeText(value).replace(/\/+$/, "");
  }

  function normalizeUrl(value) {
    const url = normalizeText(value);
    if (!url) {
      return "";
    }
    if (url.startsWith("//")) {
      return `https:${url}`;
    }
    return url.replace(/^http:\/\//i, "https://");
  }

  function normalizeRole(role) {
    if (role === "seller" || role === "viewer" || role === "admin") {
      return role;
    }
    return "viewer";
  }

  function normalizeItemType(itemType) {
    const value = normalizeText(itemType).toLowerCase();
    if (value === "book" || value === "goods" || value === "bicycle" || value === "other") {
      return value;
    }
    return "other";
  }

  function defaultCategory(itemType) {
    if (itemType === "book") return "書籍";
    if (itemType === "goods") return "小物";
    if (itemType === "bicycle") return "自転車";
    return "その他";
  }

  function normalizeRoomNumber(value) {
    return String(value || "").replace(/[^\d]/g, "").slice(0, 3);
  }

  function normalizeJan(value) {
    return String(value || "").replace(/[^\d]/g, "");
  }

  function isJanBookCode(value) {
    const jan = normalizeJan(value);
    if (!/^\d{13}$/.test(jan)) {
      return false;
    }
    if (!jan.startsWith("978") && !jan.startsWith("979")) {
      return false;
    }

    const weighted = jan
      .slice(0, 12)
      .split("")
      .map((digit) => Number(digit))
      .reduce((sum, digit, index) => sum + digit * (index % 2 === 0 ? 1 : 3), 0);
    const check = (10 - (weighted % 10)) % 10;
    return check === Number(jan.slice(-1));
  }

  function uniqueStrings(values) {
    const out = [];
    const seen = new Set();
    (values || []).forEach((value) => {
      const text = normalizeText(value);
      if (!text || seen.has(text)) {
        return;
      }
      seen.add(text);
      out.push(text);
    });
    return out;
  }

  function wait(ms) {
    return new Promise((resolve) => window.setTimeout(resolve, ms));
  }

  function getBookCacheKey(jan) {
    return `${BOOK_CACHE_PREFIX}:${jan}`;
  }

  function readBookCache(jan) {
    try {
      const raw = window.localStorage.getItem(getBookCacheKey(jan));
      if (!raw) {
        return null;
      }
      const parsed = JSON.parse(raw);
      if (!parsed?.savedAt || Date.now() - parsed.savedAt > BOOK_CACHE_TTL_MS) {
        window.localStorage.removeItem(getBookCacheKey(jan));
        return null;
      }
      return sanitizeBookMeta(parsed.data || {}, jan);
    } catch {
      return null;
    }
  }

  function writeBookCache(jan, meta) {
    try {
      window.localStorage.setItem(
        getBookCacheKey(jan),
        JSON.stringify({
          savedAt: Date.now(),
          data: sanitizeBookMeta(meta, jan),
        })
      );
    } catch {
      // ignore quota errors
    }
  }

  function sanitizeBookMeta(rawMeta, jan) {
    const isbn = normalizeJan(jan);
    const coverCandidates = uniqueStrings([
      normalizeUrl(rawMeta.coverUrl),
      ...(Array.isArray(rawMeta.coverCandidates) ? rawMeta.coverCandidates.map(normalizeUrl) : []),
      `https://cover.openbd.jp/${isbn}.jpg`,
      `https://books.google.com/books/content?vid=ISBN${isbn}&printsec=frontcover&img=1&zoom=1&source=gbs_api`,
      `https://covers.openlibrary.org/b/isbn/${isbn}-L.jpg`,
      `https://covers.openlibrary.org/b/isbn/${isbn}-M.jpg`,
    ]);

    return {
      jan: isbn,
      title: normalizeText(rawMeta.title),
      summary: normalizeText(rawMeta.summary),
      coverUrl: normalizeUrl(rawMeta.coverUrl) || coverCandidates[0] || "",
      coverCandidates,
      author: normalizeText(rawMeta.author),
      publisher: normalizeText(rawMeta.publisher),
      publishedDate: normalizeText(rawMeta.publishedDate),
    };
  }

  function mergeBookMeta(base, patch, jan) {
    const current = sanitizeBookMeta(base || {}, jan);
    const next = sanitizeBookMeta(patch || {}, jan);

    return sanitizeBookMeta(
      {
        title: current.title || next.title,
        summary: pickLongerText(current.summary, next.summary),
        coverUrl: current.coverUrl || next.coverUrl,
        coverCandidates: uniqueStrings([...(current.coverCandidates || []), ...(next.coverCandidates || [])]),
        author: current.author || next.author,
        publisher: current.publisher || next.publisher,
        publishedDate: current.publishedDate || next.publishedDate,
      },
      jan
    );
  }

  function pickLongerText(a, b) {
    const ta = normalizeText(a);
    const tb = normalizeText(b);
    return tb.length > ta.length ? tb : ta;
  }

  function hasDetailedMeta(meta) {
    const summary = normalizeText(meta.summary);
    const cover = normalizeUrl(meta.coverUrl);
    return summary.length >= BOOK_MIN_SUMMARY_LENGTH && Boolean(cover);
  }

  async function lookupBookByJan(rawJan, options = {}) {
    const jan = normalizeJan(rawJan);
    if (!isJanBookCode(jan)) {
      throw new Error("JAN/ISBN-13 の形式が正しくありません。");
    }

    const forceRefresh = Boolean(options.forceRefresh);
    let best = sanitizeBookMeta({}, jan);

    const cached = forceRefresh ? null : readBookCache(jan);
    if (cached) {
      best = mergeBookMeta(best, cached, jan);
      if (hasDetailedMeta(best)) {
        return best;
      }
    }

    for (let i = 0; i < BOOK_TIMEOUT_STEPS.length; i += 1) {
      const timeoutMs = BOOK_TIMEOUT_STEPS[i];
      const [openBd, googleBooks, openLibrary] = await Promise.all([
        fetchOpenBd(jan, timeoutMs),
        fetchGoogleBooks(jan, timeoutMs),
        fetchOpenLibrary(jan, timeoutMs),
      ]);

      best = mergeBookMeta(best, fromOpenBd(jan, openBd), jan);
      best = mergeBookMeta(best, fromGoogleBooks(jan, googleBooks), jan);
      best = mergeBookMeta(best, fromOpenLibrary(jan, openLibrary), jan);

      if (hasDetailedMeta(best)) {
        break;
      }

      if (i < BOOK_TIMEOUT_STEPS.length - 1) {
        await wait(BOOK_RETRY_DELAY_MS);
      }
    }

    if (!normalizeText(best.summary)) {
      if (cached && normalizeText(cached.summary)) {
        best.summary = cached.summary;
      } else {
        best.summary = BOOK_SUMMARY_FALLBACK;
      }
    }

    if (!normalizeUrl(best.coverUrl) && Array.isArray(best.coverCandidates) && best.coverCandidates.length) {
      best.coverUrl = best.coverCandidates[0];
    }

    writeBookCache(jan, best);
    return best;
  }

  async function fetchOpenBd(jan, timeoutMs) {
    const json = await fetchExternalJsonWithRetry(
      `https://api.openbd.jp/v1/get?isbn=${encodeURIComponent(jan)}`,
      timeoutMs,
      BOOK_RETRY_COUNT
    );
    if (!Array.isArray(json)) {
      return null;
    }
    return json[0] || null;
  }

  async function fetchGoogleBooks(jan, timeoutMs) {
    const query = encodeURIComponent(`isbn:${jan}`);
    return await fetchExternalJsonWithRetry(
      `https://www.googleapis.com/books/v1/volumes?q=${query}&maxResults=1&projection=full&printType=books`,
      timeoutMs,
      BOOK_RETRY_COUNT
    );
  }

  async function fetchOpenLibrary(jan, timeoutMs) {
    return await fetchExternalJsonWithRetry(
      `https://openlibrary.org/api/books?bibkeys=ISBN:${encodeURIComponent(jan)}&jscmd=details&format=json`,
      timeoutMs,
      BOOK_RETRY_COUNT
    );
  }

  async function fetchExternalJsonWithRetry(url, timeoutMs, retries) {
    for (let attempt = 0; attempt <= retries; attempt += 1) {
      try {
        const response = await fetchWithTimeout(url, { method: "GET" }, timeoutMs);
        if (!response.ok) {
          const retryable = [408, 425, 429, 500, 502, 503, 504].includes(response.status);
          if (retryable && attempt < retries) {
            await wait(BOOK_RETRY_DELAY_MS * (attempt + 1));
            continue;
          }
          return null;
        }

        return await response.json();
      } catch {
        if (attempt >= retries) {
          return null;
        }
        await wait(BOOK_RETRY_DELAY_MS * (attempt + 1));
      }
    }

    return null;
  }

  function fromOpenBd(jan, openBd) {
    const summary = openBd?.summary || {};
    return {
      jan,
      title: normalizeText(summary.title),
      summary: extractOpenBdDescription(openBd),
      coverUrl: normalizeUrl(summary.cover),
      coverCandidates: [normalizeUrl(summary.cover)],
      author: normalizeText(summary.author),
      publisher: normalizeText(summary.publisher),
      publishedDate: formatDate(summary.pubdate),
    };
  }

  function fromGoogleBooks(jan, response) {
    const item = Array.isArray(response?.items) ? response.items[0] : null;
    const info = item?.volumeInfo || {};

    return {
      jan,
      title: normalizeText(info.title),
      summary: normalizeText(info.description || item?.searchInfo?.textSnippet),
      coverUrl: normalizeUrl(info?.imageLinks?.thumbnail || info?.imageLinks?.smallThumbnail),
      coverCandidates: [normalizeUrl(info?.imageLinks?.thumbnail), normalizeUrl(info?.imageLinks?.smallThumbnail)],
      author: Array.isArray(info.authors) ? info.authors.map((a) => normalizeText(a)).filter(Boolean).join("、") : "",
      publisher: normalizeText(info.publisher),
      publishedDate: formatDate(info.publishedDate),
    };
  }

  function fromOpenLibrary(jan, response) {
    const record = response?.[`ISBN:${jan}`] || {};
    const details = record.details || {};
    const covers = Array.isArray(details.covers) ? details.covers : [];
    const coverId = covers.find(Boolean);

    return {
      jan,
      title: normalizeText(details.title),
      summary: extractNestedText(details.description) || extractNestedText(details.excerpt),
      coverUrl: normalizeUrl(record.thumbnail_url),
      coverCandidates: [
        normalizeUrl(record.thumbnail_url),
        coverId ? `https://covers.openlibrary.org/b/id/${encodeURIComponent(String(coverId))}-L.jpg` : "",
      ],
      author: readOpenLibraryAuthors(details.authors),
      publisher: readOpenLibraryPublisher(details.publishers),
      publishedDate: formatDate(details.publish_date),
    };
  }

  function extractOpenBdDescription(openBd) {
    const textContent = openBd?.onix?.CollateralDetail?.TextContent;
    const rows = Array.isArray(textContent) ? textContent : textContent ? [textContent] : [];
    if (!rows.length) {
      return "";
    }

    const priority = ["03", "02", "01"];
    for (const code of priority) {
      const row = rows.find((item) => String(item?.TextType || "") === code);
      if (row) {
        const text = extractNestedText(row.Text);
        if (text) {
          return text;
        }
      }
    }

    for (const row of rows) {
      const text = extractNestedText(row.Text);
      if (text) {
        return text;
      }
    }

    return "";
  }

  function extractNestedText(value) {
    if (!value) {
      return "";
    }

    if (typeof value === "string" || typeof value === "number") {
      return normalizeText(stripHtml(String(value)));
    }

    if (Array.isArray(value)) {
      for (const item of value) {
        const text = extractNestedText(item);
        if (text) {
          return text;
        }
      }
      return "";
    }

    if (typeof value === "object") {
      const keys = ["content", "#text", "Text", "value", "name", "description"];
      for (const key of keys) {
        if (value[key] !== undefined) {
          const text = extractNestedText(value[key]);
          if (text) {
            return text;
          }
        }
      }
      for (const nested of Object.values(value)) {
        const text = extractNestedText(nested);
        if (text) {
          return text;
        }
      }
    }

    return "";
  }

  function stripHtml(text) {
    const raw = normalizeText(text);
    if (!raw) {
      return "";
    }
    const withoutTags = raw.replace(/<[^>]+>/g, " ");
    const parser = new DOMParser();
    const doc = parser.parseFromString(withoutTags, "text/html");
    return normalizeText(doc.documentElement.textContent || withoutTags);
  }

  function readOpenLibraryAuthors(authors) {
    const list = Array.isArray(authors) ? authors : [];
    const names = list
      .map((author) => {
        if (typeof author === "string") {
          return normalizeText(author);
        }
        return normalizeText(author?.name);
      })
      .filter(Boolean);
    return names.join("、");
  }

  function readOpenLibraryPublisher(publishers) {
    const list = Array.isArray(publishers) ? publishers : [];
    if (!list.length) {
      return "";
    }

    const first = list[0];
    if (typeof first === "string") {
      return normalizeText(first);
    }
    return normalizeText(first?.name) || extractNestedText(first);
  }

  function formatDate(value) {
    const text = normalizeText(value);
    if (!text) {
      return "";
    }
    const digits = text.replace(/[^\d]/g, "");
    if (digits.length === 8) {
      return `${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`;
    }
    if (digits.length === 6) {
      return `${digits.slice(0, 4)}-${digits.slice(4, 6)}`;
    }
    if (digits.length === 4) {
      return digits;
    }
    return text;
  }
})();
