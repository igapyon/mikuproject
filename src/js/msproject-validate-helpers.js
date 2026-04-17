(() => {
    function parseDateValue(value) {
        if (!value) {
            return null;
        }
        const timestamp = Date.parse(value);
        return Number.isFinite(timestamp) ? timestamp : null;
    }
    function isPlaceholderUid(value) {
        return String(value || "").trim() === "0";
    }
    function isUnassignedResourceUid(value) {
        return String(value || "").trim() === "-65535";
    }
    function describeTask(task) {
        return `UID=${task.uid}${task.name ? ` (${task.name})` : ""}`;
    }
    function isComparableOutlineNumber(value) {
        if (!value) {
            return false;
        }
        return value.split(".").every((part) => /^\d+$/.test(part));
    }
    function compareOutlineNumbers(left, right) {
        if (!left || !right) {
            return 0;
        }
        const leftParts = left.split(".").map((part) => Number(part));
        const rightParts = right.split(".").map((part) => Number(part));
        const maxLength = Math.max(leftParts.length, rightParts.length);
        for (let index = 0; index < maxLength; index += 1) {
            const leftPart = leftParts[index];
            const rightPart = rightParts[index];
            if (leftPart === undefined) {
                return -1;
            }
            if (rightPart === undefined) {
                return 1;
            }
            if (leftPart !== rightPart) {
                return leftPart - rightPart;
            }
        }
        return 0;
    }
    function detectTaskOrderIssue(tasks) {
        let previousComparableTask = null;
        for (const task of tasks) {
            if (isPlaceholderUid(task.uid)) {
                continue;
            }
            if (!isComparableOutlineNumber(task.outlineNumber)) {
                continue;
            }
            if (previousComparableTask && compareOutlineNumbers(previousComparableTask.outlineNumber, task.outlineNumber) >= 0) {
                return {
                    previous: previousComparableTask,
                    current: task
                };
            }
            previousComparableTask = task;
        }
        return null;
    }
    function describeResource(resource) {
        return `UID=${resource.uid || "(なし)"}${resource.name ? ` (${resource.name})` : ""}`;
    }
    function describeCalendar(calendar) {
        return `UID=${calendar.uid}${calendar.name ? ` (${calendar.name})` : ""}`;
    }
    function describeAssignment(assignment) {
        return `UID=${assignment.uid || "(なし)"}`;
    }
    function describeTaskRef(model, taskUid) {
        if (!taskUid) {
            return "TaskUID=(なし)";
        }
        const task = model.tasks.find((item) => item.uid === taskUid);
        return task ? `TaskUID=${taskUid}${task.name ? ` (${task.name})` : ""}` : `TaskUID=${taskUid}`;
    }
    function describeResourceRef(model, resourceUid) {
        if (!resourceUid) {
            return "ResourceUID=(なし)";
        }
        const resource = model.resources.find((item) => item.uid === resourceUid);
        return resource ? `ResourceUID=${resourceUid}${resource.name ? ` (${resource.name})` : ""}` : `ResourceUID=${resourceUid}`;
    }
    globalThis.__mikuprojectMsprojectValidateHelpers = {
        parseDateValue,
        isPlaceholderUid,
        isUnassignedResourceUid,
        describeTask,
        detectTaskOrderIssue,
        describeResource,
        describeCalendar,
        describeAssignment,
        describeTaskRef,
        describeResourceRef
    };
})();
