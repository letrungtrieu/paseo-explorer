import React, { useState } from "react";
import { Platform, ScrollView, View, type LayoutChangeEvent } from "react-native";
import type { SpreadsheetGeometry } from "./spreadsheet-geometry";
import type { useSpreadsheetScroll } from "./use-spreadsheet-scroll";

interface SpreadsheetViewportProps {
  geometry: SpreadsheetGeometry;
  scrolling: ReturnType<typeof useSpreadsheetScroll>;
  children: React.ReactNode;
}

export function SpreadsheetViewport({ geometry, scrolling, children }: SpreadsheetViewportProps) {
  const [nativeViewport, setNativeViewport] = useState({ width: 0, height: 0 });
  if (Platform.OS === "web") {
    // One two-axis scroll container keeps both scrollbars at the visible viewport edges, even on very tall sheets.
    return <ScrollView ref={(scrollView) => {
      scrolling.verticalScroll.current = scrollView;
      scrolling.horizontalScroll.current = scrollView;
    }} testID="office-sheet-data-scroll" style={{ flex: 1, minWidth: 0, minHeight: 0, overflow: "scroll" }}
      contentContainerStyle={{ width: geometry.width, height: geometry.height, minWidth: "100%", minHeight: "100%" }}
      onScroll={(event) => {
        // A single browser event carries both axes, so header synchronization and prefetch use the same measurements.
        scrolling.onVerticalScroll(event);
        scrolling.onHorizontalScroll(event);
      }} onLayout={(event) => {
        scrolling.onVerticalLayout(event);
        scrolling.onHorizontalLayout(event);
      }} scrollEventThrottle={32}>
      {children}
    </ScrollView>;
  }

  const measureNativeViewport = (event: LayoutChangeEvent) => {
    const { width, height } = event.nativeEvent.layout;
    // Native ScrollViews have one scroll axis; bound the inner viewport explicitly so it cannot grow to sheet height.
    setNativeViewport((previous) => previous.width === width && previous.height === height ? previous : { width, height });
  };
  return <View onLayout={measureNativeViewport} style={{ flex: 1, minWidth: 0, minHeight: 0, overflow: "hidden" }}>
    <ScrollView ref={scrolling.horizontalScroll} horizontal testID="office-sheet-horizontal-scroll"
      style={{ width: nativeViewport.width, height: nativeViewport.height }}
      contentContainerStyle={{ height: nativeViewport.height }}
      onScroll={scrolling.onHorizontalScroll} onLayout={scrolling.onHorizontalLayout} scrollEventThrottle={32}>
      <ScrollView ref={scrolling.verticalScroll} testID="office-sheet-vertical-scroll"
        style={{ width: Math.max(geometry.width, nativeViewport.width), height: nativeViewport.height }}
        contentContainerStyle={{ height: geometry.height }}
        onScroll={scrolling.onVerticalScroll} onLayout={scrolling.onVerticalLayout} scrollEventThrottle={32}>
        {children}
      </ScrollView>
    </ScrollView>
  </View>;
}
