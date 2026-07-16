const STORE_CONFIG = {
  koreanHouse: {
    name: "한국의집 롯데월드몰",
    shortName: "한국의집",
    spreadsheetId: "1VXAFKtm7IaK6Ns_QbrRzpP8XnzBHau6e49GXhNe0uQg"
  },

  gilchaejeong: {
    name: "길채정 압구정",
    shortName: "길채정 압구정",
    spreadsheetId: "1mmesI8_0POeqRJwcdyugLauWL0ZH4boPQn4w9YbRiT8"
  },

  soba: {
    name: "소바공방 평촌",
    shortName: "소바공방 평촌",
    spreadsheetId: "1BN0FdwGuW5KCJ9_6_gkEPQHS46MX26S5zEeB9SbsnNI"
  },

  hyojonggaeng: {
    name: "효종갱 파주",
    shortName: "효종갱 파주",
    spreadsheetId: "1Rjpso7VcbDAyvRz6-ni-OlfxVI283dZ48O39UD4dvIE"
  }
};

const DASHBOARD_CACHE_SECONDS = 600;

function doGet() {
  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle("더큰코리아 경영대시보드")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}


function getInitialData() {
  const now = new Date();

  return {
    stores: Object.keys(STORE_CONFIG).map(function(key) {
      return {
        key: key,
        name: STORE_CONFIG[key].name,
        shortName: STORE_CONFIG[key].shortName
      };
    }),

    year: Number(
      Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        "yyyy"
      )
    ),

    month: Number(
      Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        "M"
      )
    )
  };
}


function getDashboardData(
  storeKey,
  year,
  month,
  forceRefresh
) {
  year = Number(year);
  month = Number(month);

  const refresh = forceRefresh === true;

  if (!year || !month) {
    throw new Error("조회 연월이 올바르지 않습니다.");
  }

  if (storeKey === "all") {
    return getAllStoreDashboard_(
      year,
      month,
      refresh
    );
  }

  const store = STORE_CONFIG[storeKey];

  if (!store) {
    throw new Error("존재하지 않는 지점입니다.");
  }

  return getSingleStoreDashboard_(
    storeKey,
    store,
    year,
    month,
    refresh
  );
}


function getAllStoreDashboard_(
  year,
  month,
  forceRefresh
) {
  const storeList = [];
  let totalSales = 0;
  let totalPurchases = 0;
  let totalPreviousYearSales = 0;

  Object.keys(STORE_CONFIG).forEach(function(storeKey) {
    const store = STORE_CONFIG[storeKey];

    try {
      const result = getSingleStoreDashboard_(
        storeKey,
        store,
        year,
        month,
         forceRefresh
      );

      totalSales += result.summary.totalSales;
      totalPurchases += result.summary.totalPurchases;
      totalPreviousYearSales +=
        result.summary.previousYearSales || 0;

      storeList.push({
        key: storeKey,
        name: store.name,
        shortName: store.shortName,
        totalSales: result.summary.totalSales,
        totalPurchases: result.summary.totalPurchases,
        costRate: result.summary.costRate,
        averageSales: result.summary.averageSales,
        businessDays: result.summary.businessDays,
        todaySales: result.summary.todaySales,

        previousYearSales:
          result.summary.previousYearSales || 0,

        yearOnYearDifference:
          result.summary.yearOnYearDifference || 0,

        yearOnYearRate:
          result.summary.yearOnYearRate || 0,

        comparisonEndDay:
          result.summary.comparisonEndDay || 0,

        error: ""
      });

    } catch (error) {
      storeList.push({
        key: storeKey,
        name: store.name,
        shortName: store.shortName,
        totalSales: 0,
        totalPurchases: 0,
        costRate: 0,
        averageSales: 0,
        businessDays: 0,
        todaySales: 0,

         previousYearSales: 0,
         yearOnYearDifference: 0,
         yearOnYearRate: 0,
         comparisonEndDay: 0,

        error: error.message
      });
    }
  });

  return {
    mode: "all",
    storeKey: "all",
    storeName: "전체 지점",
    year: year,
    month: month,

    summary: {
      totalSales: totalSales,
      totalPurchases: totalPurchases,

      previousYearSales: totalPreviousYearSales,

      yearOnYearDifference:
       totalSales - totalPreviousYearSales,

      yearOnYearRate:
        totalPreviousYearSales > 0
          ? roundNumber_(
            (
              (totalSales - totalPreviousYearSales) /
               totalPreviousYearSales
             ) * 100,
            1
            )
            : 0,

      costRate: totalSales > 0
        ? roundNumber_((totalPurchases / totalSales) * 100, 1)
        : 0,
      averageSales: calculateAverageStoreSales_(storeList),
      businessDays: 0,
      todaySales: storeList.reduce(function(sum, store) {
        return sum + Number(store.todaySales || 0);
      }, 0)
    },

    stores: storeList,
    daily: [],
    vendors: []
  };
}


function getSingleStoreDashboard_(
  storeKey,
  store,
  year,
  month,
  forceRefresh
) {
  const cache = CacheService.getScriptCache();

  const cacheKey = buildDashboardCacheKey_(
    storeKey,
    year,
    month
  );

  if (!forceRefresh) {
    const cachedText = cache.get(cacheKey);

    if (cachedText) {
      try {
        return JSON.parse(cachedText);
      } catch (error) {
        cache.remove(cacheKey);
      }
    }
  }

  const spreadsheet = SpreadsheetApp.openById(
    store.spreadsheetId
  );

  const salesSheet = spreadsheet.getSheetByName("매출");
  const purchaseSheet = spreadsheet.getSheetByName("매입");

  if (!salesSheet) {
    throw new Error(
      store.name + " 원장에 '매출' 시트가 없습니다."
    );
  }

  if (!purchaseSheet) {
    throw new Error(
      store.name + " 원장에 '매입' 시트가 없습니다."
    );
  }

  const sales = readSalesData_(
    salesSheet,
    year,
    month
  );

  const comparisonEndDay =
  getLastSalesDay_(sales.rows);

const previousSales = readSalesData_(
  salesSheet,
  year - 1,
  month,
  comparisonEndDay
);

  const purchases = readPurchaseData_(
    purchaseSheet,
    year,
    month
  );

  const dailyMap = {};

  sales.rows.forEach(function(item) {
    const key = item.dateKey;

    if (!dailyMap[key]) {
      dailyMap[key] = {
        dateKey: key,
        dateLabel: item.dateLabel,
        sales: 0,
        purchases: 0,
        costRate: 0
      };
    }

    dailyMap[key].sales += item.amount;
  });

  purchases.rows.forEach(function(item) {
    const key = item.dateKey;

    if (!dailyMap[key]) {
      dailyMap[key] = {
        dateKey: key,
        dateLabel: item.dateLabel,
        sales: 0,
        purchases: 0,
        costRate: 0
      };
    }

    dailyMap[key].purchases += item.amount;
  });

  const daily = Object.keys(dailyMap)
    .sort()
    .map(function(key) {
      const item = dailyMap[key];

      item.costRate = item.sales > 0
        ? roundNumber_(
            (item.purchases / item.sales) * 100,
            1
          )
        : 0;

      return item;
    });

  const todayKey = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  const todayItem = dailyMap[todayKey];

  const result = {
    mode: "single",
    storeKey: storeKey,
    storeName: store.name,
    shortName: store.shortName,
    year: year,
    month: month,

    summary: {
  totalSales: sales.total,
  totalPurchases: purchases.total,

  previousYearSales: previousSales.total,

  yearOnYearDifference:
    sales.total - previousSales.total,

  yearOnYearRate:
    previousSales.total > 0
      ? roundNumber_(
          (
            (sales.total - previousSales.total) /
            previousSales.total
          ) * 100,
          1
        )
      : 0,

  comparisonEndDay: comparisonEndDay,

      costRate: sales.total > 0
        ? roundNumber_(
            (purchases.total / sales.total) * 100,
            1
          )
        : 0,

      averageSales: sales.businessDays > 0
        ? Math.round(
            sales.total / sales.businessDays
          )
        : 0,

      businessDays: sales.businessDays,

      todaySales: todayItem
        ? Number(todayItem.sales || 0)
        : 0
    },

    daily: daily,
    vendors: purchases.vendors,
    stores: []
  };

  try {
    cache.put(
      cacheKey,
      JSON.stringify(result),
      DASHBOARD_CACHE_SECONDS
    );
  } catch (error) {
    console.log(
      "캐시 저장 실패: " + error.message
    );
  }

  return result;
}


function readSalesData_(
  sheet,
  year,
  month,
  maxDay
) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      total: 0,
      businessDays: 0,
      rows: []
    };
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 4)
    .getValues();

  const rows = [];
  const businessDateMap = {};
  let total = 0;

  values.forEach(function(row) {
    const reportId = String(row[0] || "").trim();
    const rawDate = row[1];
    const amount = parseMoney_(row[2]);

    const dateInfo = getDateInfo_(
      rawDate,
      reportId
    );

    if (!dateInfo) return;

    if (
      dateInfo.year !== year ||
      dateInfo.month !== month
    ) {
      return;
    }

    if (
  maxDay &&
  dateInfo.day > Number(maxDay)
) {
  return;
}

    rows.push({
      dateKey: dateInfo.dateKey,
      dateLabel: dateInfo.dateLabel,
      amount: amount
    });

    total += amount;
    businessDateMap[dateInfo.dateKey] = true;
  });

  return {
    total: total,
    businessDays: Object.keys(businessDateMap).length,
    rows: rows
  };
}


function readPurchaseData_(sheet, year, month) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      total: 0,
      rows: [],
      vendors: []
    };
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 3)
    .getValues();

  const rows = [];
  const vendorMap = {};
  let total = 0;

  values.forEach(function(row) {
    const reportId = String(row[0] || "").trim();
    const vendorName =
      String(row[1] || "").trim() || "기타";

    const amount = parseMoney_(row[2]);

    const dateInfo = getDateInfo_(
      "",
      reportId
    );

    if (!dateInfo) return;

    if (
      dateInfo.year !== year ||
      dateInfo.month !== month
    ) {
      return;
    }

    rows.push({
      dateKey: dateInfo.dateKey,
      dateLabel: dateInfo.dateLabel,
      vendorName: vendorName,
      amount: amount
    });

    total += amount;

    if (!vendorMap[vendorName]) {
      vendorMap[vendorName] = 0;
    }

    vendorMap[vendorName] += amount;
  });

  const vendors = Object.keys(vendorMap)
    .map(function(name) {
      return {
        name: name,
        amount: vendorMap[name]
      };
    })
    .sort(function(a, b) {
      return b.amount - a.amount;
    });

  return {
    total: total,
    rows: rows,
    vendors: vendors
  };
}


function getDateInfo_(rawDate, reportId) {
  let date = null;

  if (
    Object.prototype.toString.call(rawDate) ===
      "[object Date]" &&
    !isNaN(rawDate.getTime())
  ) {
    date = rawDate;
  }

  if (!date && rawDate) {
    const dateText = String(rawDate).trim();

    const match = dateText.match(
      /(\d{4})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/
    );

    if (match) {
      date = new Date(
        Number(match[1]),
        Number(match[2]) - 1,
        Number(match[3])
      );
    }
  }

  if (!date && reportId) {
    const idMatch = String(reportId).match(
      /^(\d{2})(\d{2})(\d{2})/
    );

    if (idMatch) {
      date = new Date(
        2000 + Number(idMatch[1]),
        Number(idMatch[2]) - 1,
        Number(idMatch[3])
      );
    }
  }

  if (!date || isNaN(date.getTime())) {
    return null;
  }

  return {
    year: date.getFullYear(),
    month: date.getMonth() + 1,
    day: date.getDate(),

    dateKey: Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    ),

    dateLabel:
      (date.getMonth() + 1) +
      "월 " +
      date.getDate() +
      "일"
  };
}


function parseMoney_(value) {
  if (typeof value === "number") {
    return isNaN(value) ? 0 : value;
  }

  const cleaned = String(value || "")
    .replace(/,/g, "")
    .replace(/[^\d.-]/g, "");

  const number = Number(cleaned);

  return isNaN(number) ? 0 : number;
}


function roundNumber_(value, digits) {
  const unit = Math.pow(10, digits || 0);
  return Math.round(value * unit) / unit;
}


function calculateAverageStoreSales_(stores) {
  const activeStores = stores.filter(function(store) {
    return !store.error && store.businessDays > 0;
  });

  if (!activeStores.length) return 0;

  const sum = activeStores.reduce(function(total, store) {
    return total + Number(store.averageSales || 0);
  }, 0);

  return Math.round(sum / activeStores.length);
}
function buildDashboardCacheKey_(
  storeKey,
  year,
  month
) {
  return [
    "dashboard",
    storeKey,
    year,
    String(month).padStart(2, "0")
  ].join("_");
}
function getLastSalesDay_(rows) {
  if (!rows || !rows.length) {
    return 0;
  }

  return rows.reduce(function(maxDay, row) {
    const day = Number(
      String(row.dateKey || "").split("-")[2]
    );

    return day > maxDay ? day : maxDay;
  }, 0);
}