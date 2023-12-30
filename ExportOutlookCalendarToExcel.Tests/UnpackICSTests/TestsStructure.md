Tests here are structured by most complex element in it. As example:
- Most complex event in `PlainEvents_NoChangesAfterUnpack` is plain event. So it goes to `PlainEventsTest`.
- Most complex event in `SingleRecurringEvent_DailyRecurrence_UnpackedCorrectly` is single recurring event. So it goes to `SimpleRecurringDailyEventsTests`.
- Most complex event in `SingleRecurringEvent_DailyRecurrence_PlainEvents_UnpackedCorrectly` is single recurring event, even cause it has plain events in tests. So it goes to `SimpleRecurringDailyEventsTests`.