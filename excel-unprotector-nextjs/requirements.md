# 需求：使用 Hero UI 重构页面

## 背景

项目基于 Next.js（App Router）+ Tailwind CSS 4，已安装 `@heroui/react` 3.x。本次改造仅针对 UI 层。

## 范围

- `app/layout.tsx`、`app/page.tsx`
- `components/`：Header、Footer、Features、FeatureCard、UploadZone、ProgressContainer、ResultContainer、ErrorContainer

## 要求

1. **外观与功能一致**：布局、配色、文案及交互反馈（上传、进度、结果、错误提示）与当前页面保持一致；不引入新的视觉风格（如暗色模式、主题切换）。
2. **按收益选用 Hero UI**：不强制全量替换。仅当 Hero UI 组件能明显减少代码量或提升健壮性时使用（如 Button、Card、Progress、Tooltip）；行为高度定制的组件（如 UploadZone 的拖拽上传逻辑）可保留 Tailwind 手写实现。
3. **不改业务逻辑**：`lib/`（api、hooks、utils、types）保持不变，仅允许必要的类型联动调整。
4. **合理的附带优化**（需克制）：可访问性（语义化标签、焦点管理）、明显的响应式问题、移除替换后不再使用的样式与代码。

## 验收标准

- `npm run build` 与 `npm run lint` 通过。
- 浏览器中页面外观与交互行为与重构前一致。
