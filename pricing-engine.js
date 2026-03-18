(function (global) {
    const MS_PER_DAY = 1000 * 60 * 60 * 24;

    function addCalendarDays(date, days) {
        const result = new Date(date);
        result.setDate(result.getDate() + days);
        return result;
    }

    function diffInclusive(startDate, endDate) {
        return Math.ceil((endDate - startDate) / MS_PER_DAY) + 1;
    }

    function normalizePeriods(periods) {
        return Array.isArray(periods) ? periods : [];
    }

    function buildChargeBreakdown(chargeableDays, periods, options) {
        const safeOptions = options || {};
        const normalizedPeriods = normalizePeriods(periods);
        const firstChargedDayOffset = safeOptions.firstChargedDayOffset || 0;
        const calendarStartDate = safeOptions.calendarStartDate || null;

        let remaining = Math.max(0, chargeableDays || 0);
        let daysFromStart = firstChargedDayOffset;
        let cost = 0;
        const breakdown = [];

        for (const period of normalizedPeriods) {
            if (remaining <= 0) break;

            const periodDays = Number(period.days) || 0;
            const rate = Number(period.rate) || 0;
            const daysInPeriod = Math.min(remaining, periodDays || Infinity);
            const subtotal = daysInPeriod * rate;

            cost += subtotal;
            breakdown.push({
                startDay: daysFromStart + 1,
                endDay: daysFromStart + daysInPeriod,
                days: daysInPeriod,
                rate,
                subtotal,
                calStart: calendarStartDate ? addCalendarDays(calendarStartDate, daysFromStart) : null,
                calEnd: calendarStartDate ? addCalendarDays(calendarStartDate, daysFromStart + daysInPeriod - 1) : null,
                isLast: periodDays === 0,
            });

            remaining -= daysInPeriod;
            daysFromStart += daysInPeriod;
        }

        return { cost, breakdown };
    }

    function calculateStorage(params) {
        const releaseDate = params.releaseDate || null;
        const pickupDate = params.pickupDate || null;
        const freeDays = Math.max(0, Number(params.freeDays) || 0);
        const periods = normalizePeriods(params.periods);

        if (!releaseDate || !pickupDate) {
            return { error: 'Datas de descarregamento/levantamento em falta.' };
        }

        const totalDays = diffInclusive(releaseDate, pickupDate);
        if (totalDays <= 0) {
            return { error: 'Data de levantamento deve ser apos descarregamento.' };
        }

        const chargeableDays = Math.max(0, totalDays - freeDays);
        const chargeResult = buildChargeBreakdown(chargeableDays, periods, {
            firstChargedDayOffset: freeDays,
            calendarStartDate: releaseDate,
        });

        return {
            totalDays,
            freeDays,
            chargeableDays,
            cost: chargeResult.cost,
            breakdown: chargeResult.breakdown,
            freeStartDate: freeDays > 0 ? releaseDate : null,
            freeEndDate: freeDays > 0 ? addCalendarDays(releaseDate, freeDays - 1) : null,
        };
    }

    function calculateDetention(params) {
        const releaseDate = params.releaseDate || null;
        const pickupDate = params.pickupDate || null;
        const returnDate = params.returnDate || null;
        const detentionFreeDays = Math.max(0, Number(params.detentionFreeDays) || 0);
        const startMode = params.startMode === 'pickup' ? 'pickup' : 'release';
        const periods = normalizePeriods(params.periods);

        if (!returnDate) return null;

        const startDate = startMode === 'release' ? releaseDate : pickupDate;
        if (!startDate) return null;

        const totalDays = diffInclusive(startDate, returnDate);
        if (totalDays <= 0) return null;

        const chargeableDays = Math.max(0, totalDays - detentionFreeDays);
        const chargeResult = buildChargeBreakdown(chargeableDays, periods, {
            firstChargedDayOffset: detentionFreeDays,
            calendarStartDate: startDate,
        });

        return {
            totalDays,
            detentionFreeDays,
            chargeableDays,
            cost: chargeResult.cost,
            breakdown: chargeResult.breakdown,
            freeStartDate: detentionFreeDays > 0 ? startDate : null,
            freeEndDate: detentionFreeDays > 0 ? addCalendarDays(startDate, detentionFreeDays - 1) : null,
            returnDate,
            startDate,
            startMode,
        };
    }

    global.StorageDetentionPricing = {
        calculateStorage,
        calculateDetention,
    };
})(window);
