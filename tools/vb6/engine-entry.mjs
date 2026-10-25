/** Headless bundle entry for Melina Office. Re-exports the engine and the
 *  declaration scan the editor uses for completion. Not ApplicationHost. */
import { compileProject, VirtualMachine, AutomationRegistry, describe, unbox, NOTHING, MISSING, VBError, VBArray } from '../../third_party/VB6/src/runtime/engine-entry.js';
import { scanDeclarations } from '../../third_party/VB6/src/language/declaration-index.js';
import { statementBefore, expressionBefore, completionSpan, completionMatches } from '../../third_party/VB6/src/language/source-context.js';
import { intrinsicCatalog } from './intrinsic-catalog.js';

export { compileProject, VirtualMachine, AutomationRegistry, describe, unbox, NOTHING, MISSING, VBError, VBArray, scanDeclarations, statementBefore, expressionBefore, completionSpan, completionMatches, intrinsicCatalog };
