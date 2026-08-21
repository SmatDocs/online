/* -*- js-indent-level: 8 -*- */
/*
 * Copyright the Collabora Online contributors.
 *
 * SPDX-License-Identifier: MPL-2.0
 */

describe('Initial loading screen lifecycle', function () {
	it('renders tenant-neutral feedback before the browser scripts initialize', function () {
		const source = require('fs').readFileSync('html/cool.html.m4', 'utf8');

		nodeassert.match(source, /id="initial-loading-screen"/);
		nodeassert.match(source, /Initializing\.\.\./);
	});

	it('dismisses the static screen from both progress completion paths', function () {
		const source = require('fs').readFileSync(
			'src/layer/marker/ProgressOverlay.js',
			'utf8',
		);

		nodeassert.match(
			source,
			/showSpinner:[\s\S]*this\._dismissInitialLoadingScreen\(\)/,
		);
		nodeassert.match(
			source,
			/hideSpinner:[\s\S]*this\._dismissInitialLoadingScreen\(\)/,
		);
		nodeassert.match(source, /getElementById\('initial-loading-screen'\)/);
		nodeassert.doesNotMatch(source, /brandProductName/);
		nodeassert.doesNotMatch(source, /brand-label/);
	});
});
