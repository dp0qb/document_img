local ensure_packer = function()
  local fn = vim.fn
  local install_path = fn.stdpath('data')..'/site/pack/packer/start/packer.nvim'
  if fn.empty(fn.glob(install_path)) > 0 then
    fn.system({'git', 'clone', '--depth', '1', 'https://github.com/wbthomason/packer.nvim', install_path})
    vim.cmd [[packadd packer.nvim]]
    return true
  end
  return false
end

local packer_bootstrap = ensure_packer()

vim.cmd([[
  augroup packer_user_config
    autocmd!
    autocmd BufWritePost plugins-setup.lua source <afile> | PackerSync
  augroup end
]])

return require('packer').startup(function(use)
  use 'wbthomason/packer.nvim'

  -- theme start
  use 'folke/tokyonight.nvim'
  -- theme end

  -- status-line start
  use {
    'nvim-lualine/lualine.nvim',
    requires = { 'kyazdani42/nvim-web-devicons', opt = true }
  }
  -- status-line end

  -- navigation start
  use {
  'nvim-tree/nvim-tree.lua',
    requires = {
      'nvim-tree/nvim-web-devicons', -- optional
    }
  }
  use 'christoomey/vim-tmux-navigator'
  use('akinsho/bufferline.nvim')
  use('lewis6991/gitsigns.nvim')

  -- navigation end

  -- highlight end
  use 'nvim-treesitter/nvim-treesitter'
  use 'p00f/nvim-ts-rainbow'
  use("windwp/nvim-autopairs")
  use("numToStr/Comment.nvim")
  -- highlight end

  -- telescope start
  use("nvim-lua/popup.nvim")
  use("nvim-lua/plenary.nvim")
  use {
    'nvim-telescope/telescope.nvim', tag = '0.1.1',
-- or                            , branch = '0.1.x',
    requires = { {'nvim-lua/plenary.nvim'} }
  }
  use({
    "j-hui/fidget.nvim",
     config = function()
       require("fidget").setup()
     end
  })
  -- telescope end

  -- lsp and mason start
  use('williamboman/mason.nvim')
  use('williamboman/mason-lspconfig.nvim')
  use('neovim/nvim-lspconfig')
  -- lsp and mason end

  -- cmp start
  use('hrsh7th/nvim-cmp')
  use('hrsh7th/cmp-nvim-lsp')
  use('hrsh7th/cmp-vsnip')
  use('hrsh7th/cmp-path')
  use('hrsh7th/cmp-buffer')
  use('hrsh7th/vim-vsnip')
  use('rafamadriz/friendly-snippets')
  use('saadparwaiz1/cmp_luasnip')
  -- cmp end

  -- debug start
  use 'puremourning/vimspector'
  use 'mfussenegger/nvim-dap'
  -- debug end

  -- rust start
  use('simrat39/rust-tools.nvim')
  -- rust end

  if packer_bootstrap then
    require('packer').sync()
  end
end)