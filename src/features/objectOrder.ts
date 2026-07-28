/**
 * Group, ungroup, and z-order controls. Not a VBA port like the rest of
 * this app — the original toolbar's icon set (assets/icons/, Phase 8c)
 * has nothing matching these concepts, since native PowerPoint already
 * exposes Group/Ungroup/Bring to Front/etc. on its own ribbon. Added here
 * purely for convenience within this add-in's own panel.
 *
 * addGroup, ShapeGroup.ungroup, and Shape.setZOrder are all real,
 * confirmed PowerPointApi 1.8 — the same requirement-set floor
 * libraryInsert.ts already assumes for its own addGroup call.
 */

export function isGroupingSupported(): boolean {
  return Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
}

/** Groups the selected shapes into one. Requires 2+ shapes selected. */
export async function groupShapes(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    if (shapes.items.length < 2) {
      throw new Error("Select at least 2 objects and try again.");
    }
    // Children must exist server-side before addGroup can reference them —
    // already true here since they're pre-existing selected shapes, not
    // freshly built ones (unlike libraryInsert.ts's reconstruct path).
    context.presentation.getSelectedSlides().getItemAt(0).shapes.addGroup(shapes.items);
    await context.sync();
  });
}

/** Ungroups every grouped shape in the selection; non-group shapes in the same selection are left alone. */
export async function ungroupShapes(): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    const groups = shapes.items.filter((s) => s.type === PowerPoint.ShapeType.group);
    if (groups.length === 0) {
      throw new Error("Select a grouped object to ungroup.");
    }
    for (const shape of groups) {
      shape.group.ungroup();
    }
    await context.sync();
  });
}

async function setZOrderForSelection(position: PowerPoint.ShapeZOrder): Promise<void> {
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items");
    await context.sync();
    if (shapes.items.length === 0) {
      throw new Error("Select one or more objects first.");
    }
    for (const shape of shapes.items) {
      shape.setZOrder(position);
    }
    await context.sync();
  });
}

export async function bringToFront(): Promise<void> {
  await setZOrderForSelection(PowerPoint.ShapeZOrder.bringToFront);
}

export async function bringForward(): Promise<void> {
  await setZOrderForSelection(PowerPoint.ShapeZOrder.bringForward);
}

export async function sendBackward(): Promise<void> {
  await setZOrderForSelection(PowerPoint.ShapeZOrder.sendBackward);
}

export async function sendToBack(): Promise<void> {
  await setZOrderForSelection(PowerPoint.ShapeZOrder.sendToBack);
}
