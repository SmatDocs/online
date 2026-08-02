/* -*- js-indent-level: 8 -*- */
/*
 * Copyright the Collabora Online contributors.
 *
 * SPDX-License-Identifier: MPL-2.0
 */

describe('Sidebar initial visibility', function() {
	it('keeps automatic startup deck payloads hidden', function() {
		let showCalls = 0;
		const sidebar = {
			shouldShowSidebar: Sidebar.prototype.shouldShowSidebar,
			isVisible() {
				return false;
			},
			showSidebar() {
				showCalls++;
			},
			closeSidebar() {
				throw new Error('an already hidden sidebar must not be closed again');
			},
			sidebarShownTheFirstTime: true,
			isUserRequest: false,
			map: {
				uiManager: {
					getBooleanDocTypePref(name: string, defaultValue: boolean) {
						nodeassert.strictEqual(name, 'ShowSidebar');
						nodeassert.strictEqual(defaultValue, false);
						return false;
					},
				},
			},
		};

		const shouldShow = Sidebar.prototype.applySidebarVisibility.call(sidebar);

		nodeassert.strictEqual(shouldShow, false);
		nodeassert.strictEqual(showCalls, 0);
	});

	it('allows an explicit open request after startup', function() {
		let showCalls = 0;
		const sidebar = {
			shouldShowSidebar: Sidebar.prototype.shouldShowSidebar,
			isVisible() {
				return false;
			},
			showSidebar() {
				showCalls++;
			},
			closeSidebar() {
				throw new Error('an explicit open must not close the sidebar');
			},
			sidebarShownTheFirstTime: false,
			isUserRequest: false,
			map: {
				uiManager: {
					getBooleanDocTypePref() {
						return true;
					},
				},
			},
		};

		const shouldShow = Sidebar.prototype.applySidebarVisibility.call(sidebar);

		nodeassert.strictEqual(shouldShow, true);
		nodeassert.strictEqual(showCalls, 1);
		nodeassert.strictEqual(sidebar.isUserRequest, true);
	});
});
