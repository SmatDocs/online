/* -*- js-indent-level: 8 -*- */
/*
 * Copyright the Collabora Online contributors.
 *
 * SPDX-License-Identifier: MPL-2.0
 */

describe('Rename interaction', function() {
	let canvas: HTMLCanvasElement;

	beforeEach(function() {
		canvas = document.createElement('canvas');
		canvas.id = 'document-canvas';
		canvas.style.pointerEvents = 'auto';
		document.body.appendChild(canvas);
	});

	afterEach(function() {
		canvas.remove();
	});

	it('keeps navigation pointer events available while mutations are blocked', function() {
		let busyCalls = 0;
		let loadingAnimationCalls = 0;
		const uiManager = {
			blockedUI: false,
			map: {
				fire(eventName: string) {
					if (eventName === 'showbusy') busyCalls++;
				},
			},
			documentNameInput: {
				showLoadingAnimation() {
					loadingAnimationCalls++;
				},
			},
			beginRenameInteraction: UIManager.prototype.beginRenameInteraction,
		};

		UIManager.prototype.blockUI.call(uiManager, { reason: 'rename' });

		nodeassert.strictEqual(uiManager.blockedUI, true);
		nodeassert.strictEqual(loadingAnimationCalls, 1);
		nodeassert.strictEqual(busyCalls, 0);
		nodeassert.strictEqual(canvas.style.pointerEvents, 'auto');
	});

	it('retains the modal pointer lock for non-rename operations', function() {
		let busyCalls = 0;
		const uiManager = {
			blockedUI: false,
			map: {
				fire(eventName: string) {
					if (eventName === 'showbusy') busyCalls++;
				},
			},
		};

		UIManager.prototype.blockUI.call(uiManager, { reason: 'switchingtooffline' });

		nodeassert.strictEqual(uiManager.blockedUI, true);
		nodeassert.strictEqual(busyCalls, 1);
	});
});
